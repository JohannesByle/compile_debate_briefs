from googleapiclient import discovery, http, errors
from httplib2 import Http
import os
from oauth2client import file, client, tools
import certifi
from datetime import datetime
from io import StringIO
import pandas as pd
import time
from tqdm import tqdm

path = os.path.dirname(__file__)
cache_path = os.path.join(path, "pdf_data")
if not os.path.exists(cache_path):
    os.mkdir(cache_path)
storage_path = os.path.join("storage.json")
client_id_path = os.path.join(path, "client_id.json")
if not os.path.exists(client_id_path):
    error_url = "https://developers.google.com/drive/api/v3/about-auth"
    raise Exception(f"Client id not found, please create a client_id.json file. For more info visit {error_url}")


def build_drive():
    scopes = 'https://www.googleapis.com/auth/drive'
    store = file.Storage(storage_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(client_id_path, scopes)
        credentials = tools.run_flow(flow, store, http=Http(ca_certs=certifi.where()))
    drive_ = discovery.build('drive', 'v3', http=credentials.authorize(Http(ca_certs=certifi.where())))
    return drive_


def download_brief(file_id, drive_=None, max_tries=10):
    # Build drive if not already created
    if not drive_:
        drive_ = build_drive()

    # Get file meta-data
    try:
        file_ = drive_.files().get(fileId=file_id, fields='parents, mimeType, modifiedTime').execute()
    except errors.HttpError as e:
        if e.resp.status == 404:
            return "The file is not shared with the debate team"
        elif e.resp.status == 503 and max_tries > 0:
            time.sleep(1)
            return download_brief(file_id, drive_=drive_, max_tries=max_tries - 1)
        else:
            return str(e)
    drive_time = datetime.strptime(file_["modifiedTime"], '%Y-%m-%dT%H:%M:%S.%fZ').timestamp()

    # Check file type and download accordingly if .pdf does not already exist
    file_path = os.path.join(cache_path, file_id)
    if os.path.exists(file_path + ".pdf") and drive_time < os.path.getmtime(file_path + ".pdf"):
        return
    if file_['mimeType'] == 'application/vnd.google-apps.document':
        mimetype = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        data = drive_.files().export(fileId=file_id, mimeType=mimetype).execute()
    elif file_['mimeType'] == 'application/vnd.openxmlformats-officedocument.wordprocessingml.document':
        data = drive_.files().get_media(fileId=file_id).execute()
    else:
        raise Exception("Incorrect format, the requested file is not a google doc")

    # Save downloaded file as .docx
    with open(os.path.join(cache_path, file_id + '.docx'), 'wb') as f:
        f.write(data)

    # Convert .docx to pdf that is formatted according to LaTex conventions
    args = "-V geometry:margin=0.6in -V geometry:bottom=1in -V pagestyle=empty"
    os.system(f"pandoc {args} -o {file_path}.pdf {file_path}.docx --pdf-engine=xelatex")
    if not os.path.exists(file_path + ".pdf"):
        raise Exception("Converting file failed")

    # Share file with debate team
    briefs_folder_id = "1LJ8aOEvfg6Q1jgNjto_mdN2u1rVkhWc5"
    if "parents" in file_ and briefs_folder_id not in file_["parents"]:
        drive_.files().update(fileId=file_id, addParents=briefs_folder_id, fields='id, parents').execute()

    # Delete .docx file
    if os.path.exists(file_path + '.docx'):
        os.remove(file_path + ".docx")


def get_briefs():
    drive = build_drive()

    # Get spreadsheet
    file_id = "1PstwVA00z1YY-3FAcQcfgmM83SWUe2jR7e2I4zEhThM"
    sheet = drive.files().export(fileId=file_id, mimeType='text/csv').execute().decode('utf-8')
    df = pd.read_csv(StringIO(sheet))
    df = df.dropna(subset=["Title"])
    df["Date"] = pd.to_datetime(df["Date"], format="%m/%d/%Y")
    df = df.sort_values("Date")
    df["Title"] = df["Title"].str.strip()
    df["file_id"] = df["Link"].str.extract(r"d/(.+)/")
    df.index = df.index.astype(str)

    # Download briefs from spreadsheet
    p_bar = tqdm(df.sort_index().iterrows(), total=len(df))
    for index, row in p_bar:
        p_bar.set_description(row["Title"])
        if isinstance(row["Link"], str) and isinstance(row["Title"], str):
            try:
                error = download_brief(row["file_id"])
            except Exception as e:
                error = str(e)
            if error:
                df = df.drop(index)
    return df, drive


def get_pdf(drive_=None):
    if not drive_:
        drive_ = build_drive()
    # Get list of all files in Google (Fuck Google) Drive
    directory = drive_.files().list(fields='files(description, id, mimeType)').execute()

    # Loop over files in directory and look for old indexed briefs .pdf file
    for files in directory.get('files'):

        # If file matches description of indexed_briefs delete it
        try:
            if files['description'] == 'GlckOayFQgdIdOqRBOL8' and files['mimeType'] == 'application/pdf':
                return files
        except Exception:
            pass
    return None
