from __future__ import print_function
from googleapiclient import discovery, http
from httplib2 import Http
from oauth2client import file, client, tools
import re
from os import remove, system
import os
from datetime import datetime, date
from tqdm import tqdm
import pandas as pd
from io import StringIO

# Initialize the api
path = os.path.dirname(__file__)
SCOPES = 'https://www.googleapis.com/auth/drive'
store = file.Storage(os.path.join(path, 'storage.json'))
creds = store.get()
if not creds or creds.invalid:
    flow = client.flow_from_clientsecrets(os.path.join(path, 'client_id.json'), SCOPES)
    creds = tools.run_flow(flow, store, http=Http(disable_ssl_certificate_validation=True))
DRIVE = discovery.build('drive', 'v3', http=creds.authorize(Http(disable_ssl_certificate_validation=True)))

# Open the spreadsheet
file_id = '1_OlTo22DVtVYwUQM5EAx9JhWO5a6_-O6y8Q32nV6KP8'
sheet = DRIVE.files().export(fileId=file_id, mimeType='text/csv').execute().decode('utf-8')
df = pd.read_csv(StringIO(sheet))
df["Date"] = pd.to_datetime(df["Date"], format="%m/%d/%Y")
df = df.sort_values("Date")
df["Title"] = df["Title"].str.strip()
df["file_id"] = df["Link"].str.extract(r"d/(.+)/")


def download_brief(file_id):
    # Get file meta-data
    file = DRIVE.files().get(fileId=file_id, fields='parents, mimeType, modifiedTime').execute()
    drive_time = datetime.strptime(file["modifiedTime"], '%Y-%m-%dT%H:%M:%S.%fZ').timestamp()

    # Check file type and download accordingly if .pdf does not already exist
    if not os.path.exists("pdf_data/" + file_id + '.pdf') or drive_time > os.path.getmtime(
            "pdf_data/" + file_id + '.pdf'):

        if file['mimeType'] == 'application/vnd.google-apps.document':
            MIMETYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            data = DRIVE.files().export(fileId=file_id, mimeType=MIMETYPE).execute()
        elif file['mimeType'] == 'application/vnd.openxmlformats-officedocument.wordprocessingml.document':
            data = DRIVE.files().get_media(fileId=file_id).execute()
        else:
            print(file_id + " has the wrong format and will not be included. Please fix this file")
            return

        # Save downloaded file as .docx
        with open("pdf_data/" + file_id + '.docx', 'wb') as f:
            f.write(data)

        # Convert .docx to pdf that is formatted according to LaTex conventions
        system(r"pandoc -V geometry:margin=0.6in -V geometry:bottom=1in -V pagestyle=empty --quiet -o " +
               "pdf_data/" + file_id + r".pdf " + "pdf_data/" + file_id + r".docx --pdf-engine=xelatex")

    # Share file with debate team
    if "1LJ8aOEvfg6Q1jgNjto_mdN2u1rVkhWc5" not in file["parents"]:
        DRIVE.files().update(fileId=file_id,
                             addParents="1LJ8aOEvfg6Q1jgNjto_mdN2u1rVkhWc5",
                             fields='id, parents').execute()

    # Delete .docx file
    if os.path.exists("pdf_data/" + file_id + '.docx'):
        remove("pdf_data/" + file_id + ".docx")


pbar = tqdm(df.iterrows(), total=len(df.index))
for index, row in pbar:
    if isinstance(row["Link"], str) and isinstance(row["Title"], str):
        pbar.set_description(row["Title"])
        try:
            download_brief(row["file_id"])
        except Exception as e:
            print()
            print("An error occurred while scraping " + row["Title"])
            df = df.drop(index)

with open("template.tex", "r") as f:
    tex_string = f.read()


# Generate category list
categories = pd.unique(df[['Categories', 'Unnamed: 6', 'Unnamed: 7', 'Unnamed: 8', 'Unnamed: 9']].values.ravel("K"))
categories = [n for n in categories if isinstance(n, str)]
categories_latex_list = ""
for category in sorted(categories):
    categories_latex_list += r"\subsection*{@}".replace("@", category) + "\n"
    briefs = df[(df['Categories'] == category) | (df['Unnamed: 6'] == category) | (
                df['Unnamed: 7'] == category) | (df['Unnamed: 8'] == category) | (df['Unnamed: 9'] == category)]
    for index, brief in briefs.iterrows():
        categories_latex_list += "\t" + brief["Title"] + r"\dotfill @\\".replace("@", str(index)) + "\n"

tex_string = tex_string.replace("% Table of Contents - 1", categories_latex_list)

# Generate flat list
flat_latex_list = ""
pages_latex_list = ""
page_template = r"\label{@id}\fancyhead[C]{@index. @title}\includepdf[pages = -, pagecommand = {}]{pdf_data/@id.pdf}"
for index, brief in df.sort_values("Title").iterrows():
    if isinstance(brief["Title"], str):
        flat_latex_list += brief["Title"] + r"\dotfill @\\".replace("@", str(index)) + "\n"

tex_string = tex_string.replace("% Table of Contents - 2", flat_latex_list)

# Generate pages list
for index, brief in df.sort_index().iterrows():
    if isinstance(brief["Title"], str):
        pages_latex_list += page_template.replace("@id", brief["file_id"]).replace("@index", str(index)). \
                            replace("@title", brief["Title"]) + "\n"

tex_string = tex_string.replace("% Pages", pages_latex_list)

if os.path.exists("indexed_briefs.aux"):
    os.remove("indexed_briefs.aux")

with open("indexed_briefs.tex", "w") as f:
    f.write(tex_string)

system(r"pdflatex indexed_briefs.tex")

# Get list of all files in Google (Fuck Google) Drive
directory = DRIVE.files().list(fields='files(description, id, mimeType)').execute()

# Loop over files in directory and look for old indexed briefs .pdf file
for files in directory.get('files'):

    # If file matches description of indexed_briefs delete it
    try:
        if files['description'] == 'GlckOayFQgdIdOqRBOL8' and files['mimeType'] == 'application/pdf':
            removed = DRIVE.files().delete(fileId=files['id']).execute()
    except Exception:
        pass

# Upload new file
file_metadata = {'name': 'Indexed Briefs (' + str(date.today()) + ')',
                 'description': 'GlckOayFQgdIdOqRBOL8',
                 "parents": ['1PSgntCxfM-2YidrIjS8hzfzdzoDGv0ze']}
media = http.MediaFileUpload('indexed_briefs.pdf', mimetype='application/pdf')
file = DRIVE.files().create(body=file_metadata, media_body=media).execute()
