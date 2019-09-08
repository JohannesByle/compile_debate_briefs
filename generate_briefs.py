# Author: Johannes Byle
from __future__ import print_function
from googleapiclient import discovery, http
from httplib2 import Http
from oauth2client import file, client, tools
import re
from os import path, remove, system, _exit
import sys
import csv
import datetime
from tqdm import tqdm

# Initialize the necessary variable for googleapiclient to work
SCOPES = ['https://www.googleapis.com/auth/drive']
store = file.Storage('storage.json')
creds = store.get()

# Check if script is being run for the first time, in which case it will ask you to sign in again
# This feature might not be necessary and might need to be removed
if not path.exists('auth.txt'):
    flow = client.flow_from_clientsecrets('client_id.json', SCOPES)
    creds = tools.run_flow(flow, store, http=creds.authorize(Http(disable_ssl_certificate_validation=True)))
    with open('auth.txt', 'w') as f:
        f.write("Google")

# Initialize the main Google variable
# noinspection PyBroadException
try:
    DRIVE = discovery.build('drive', 'v3', http=creds.authorize(Http(disable_ssl_certificate_validation=True)))
except Exception:
    print("Something went wrong while trying to talk with Google")

    # If something is wrong with the credentials it might be fixed by signing in again, in which case delete auth file
    # so that login code can br run on the next execution
    # noinspection PyBroadException
    try:
        remove('auth.txt')
    except Exception:
        pass

    sys.exit()

# Open the spreadsheet
file_id = '1_OlTo22DVtVYwUQM5EAx9JhWO5a6_-O6y8Q32nV6KP8'
sheet = csv.reader(str.splitlines(DRIVE.files().export(fileId=file_id, mimeType='text/csv').execute().decode('utf-8')))

# Initialize the variables necessary to extract data from sheet
categories = []
tags = []
titles = []
dates = []
sheet_data = []

# Remove first two lines from sheet and convert to multidimensional list
for row in sheet:
    sheet_data.append(row)
sheet_data = sheet_data[2:]

# Extract data from sheet
for row in sheet_data:

    # If category is not already in list add it to list
    for category in row[5:10]:
        if category not in categories and category != '':
            categories.append(category)

    # Create list of tags, titles and dates
    if row[2].startswith('https'):
        tags.append(re.search(r'd/(.+)/', str(row[2])).group(1))
        titles.append(row[3])
        dates.append(datetime.datetime.strptime(row[4], '%m/%d/%Y'))

# Sort tags and titles by date
dates, titles, tags = zip(*sorted(zip(dates, titles, tags)))


# Initialize variables for storing titles and tags by category
categories = sorted(categories)
titles_by_category = categories.copy()
tags_by_category = categories.copy()

# Loop through categories and append tags and titles corresponding to those categories
for category in range(len(categories)):
    titles_by_category[category] = []
    tags_by_category[category] = []
    for row in sheet_data:
        if categories[category] in row[5:10]:
            titles_by_category[category].append(row[3])
            tags_by_category[category].append(tags[titles.index(row[3])])

# Generate starting LaTex code
latex_pre = r'''
\documentclass{article}
\usepackage[margin=0.6in, bottom=0.9in]{geometry}
\usepackage{pdfpages}
\usepackage{multicol}
\begin{document}
\begin{multicols}{2}
\pagenumbering{roman}
'''

# Generate ending LaTex code
latex_post = r'''
\end{document}
'''

# Generate title for index of categories
category_index = r'''
\end{multicols}
\begin{center}
\section*{Index of Briefs by Category}
\end{center}
\begin{multicols}{2}
\noindent
'''

# Generate index of briefs by category
for category in range(len(categories)):
    category_index = category_index + r'''
    \subsection*{''' + categories[category] + '''}
        '''
    for n in range(len(titles_by_category[category])):
        category_index = category_index + titles_by_category[category][n] + r'''
        \dotfill
        \pageref{''' + tags_by_category[category][n] + r'''}\\
        '''

# Generate title of list of briefs
briefs_list = r'''
\end{multicols}
\begin{center}
\section*{List of All Briefs}
\end{center}
\begin{multicols}{2}
\noindent
    '''

# Generate list of briefs
for n in range(len(titles)):
    briefs_list = briefs_list + titles[n] + r'''
    \dotfill
    \pageref{''' + tags[n] + r'''}\\
    '''

# Generate content
content = r'''
\end{multicols}
\pagebreak
\pagenumbering{arabic}
'''


# Loop over tags and download file referenced by tag
pbar = tqdm(tags)
for file_id in pbar:

    # Try downloading files from Google drive
    # noinspection PyBroadException
    try:

        # Get file meta-data
        title_file = DRIVE.files().get(fileId=file_id, fields='name, parents').execute()
        title = title_file['name']

        # Move file to Briefs folder
        title_file = DRIVE.files().update(fileId=file_id, addParents='1LJ8aOEvfg6Q1jgNjto_mdN2u1rVkhWc5').execute()

        # Set progress bar description
        pbar.set_description("Processing: "+title)

        # Check file type and download accordingly if .pdf does not already exist
        if not path.exists(file_id+'.pdf'):

            if title_file['mimeType'] == 'application/vnd.google-apps.document':
                MIMETYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                data = DRIVE.files().export(fileId=file_id, mimeType=MIMETYPE).execute()
            elif title_file['mimeType'] == 'application/vnd.openxmlformats-officedocument.wordprocessingml.document':
                data = DRIVE.files().get_media(fileId=file_id).execute()
            else:
                print(title + " has the wrong format and will not be included. Please fix this from the spreadsheet")
                continue
            # Save downloaded file as .docx
            with open(file_id+'.docx', 'wb') as f:
                f.write(data)

            # Convert .docx to pdf that is formatted according to LaTex conventions
            system(r"pandoc -V geometry:margin=0.6in -V geometry:bottom=1in -V pagestyle=empty --quiet -o " + file_id + r".pdf " +
                   file_id + r".docx --pdf-engine=xelatex")

            # Delete .docx file
            remove(file_id+".docx")

        # Append link to pdf to LaTex code
        content = content + r'''
            \label{''' + file_id + r'''}
            \includepdf[pages=-, pagecommand={}]{''' + file_id + r'''.pdf}
            '''
    except Exception as e:
        print(e)
        print("Something went wrong with: " + titles[tags.index(file_id)])

# Combine sections into one document
final_latex = latex_pre + category_index + briefs_list + content + latex_post

# Generate the tex file
with open('indexed_briefs.tex', 'w') as f:
    f.write(final_latex)

# Generate pdf from the tex file
system(r"pdflatex indexed_briefs.tex")

# Generate pdf from the tex file a second time to make sure \pageref is working properly
system(r"pdflatex indexed_briefs.tex")

# Get list of all files in Google (Fuck Google) Drive
directory = DRIVE.files().list(fields='files(description, id, mimeType)').execute()

# Loop over files in directory and look for old indexed briefs .pdf file
for files in directory.get('files'):

    # If file matches description of indexed_briefs delete it
    # noinspection PyBroadException
    try:
        if files['description'] == 'GlckOayFQgdIdOqRBOL8' and files['mimeType'] == 'application/pdf':
            removed = DRIVE.files().delete(fileId=files['id']).execute()
    except Exception:
        pass

# Upload new file
file_metadata = {'name': 'Indexed Briefs ('+str(datetime.date.today())+')', 'description': 'GlckOayFQgdIdOqRBOL8'}
media = http.MediaFileUpload('indexed_briefs.pdf', mimetype='application/pdf')
file = DRIVE.files().create(body=file_metadata, media_body=media).execute()

# Move file to Wheaton Debate
file = DRIVE.files().update(fileId=file['id'], addParents='1PSgntCxfM-2YidrIjS8hzfzdzoDGv0ze').execute()

# Remove temporary pdf documents
for file_id in tags:
    if path.exists(file_id + ".pdf"):
        remove(file_id + ".pdf")
