import datetime
import os
import pickle
import re

import pandas
import requests
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from tqdm import tqdm

# Global variables
cwd = os.getcwd()
clientjson = 'client_secret_622133740899-o9q7o6abncdsnmrmdtnb232dq85g5h8u.apps.googleusercontent.com.json'
mode = 'test'
folder_id = '195hNR5Z9oa5-rae04xC50zRE-WRhqYgy'
# folder_id = '1-Bl9gO1RqIkXn9m1laZ04WAzcaRxHVkI'
datestring = datetime.date.today().strftime("%Y-%m-%d")


def get_gdrive_service():
    # If modifying these scopes, delete the file token.pickle.
    SCOPES = [
        "https://www.googleapis.com/auth/drive.metadata",
        "https://www.googleapis.com/auth/drive",
        "https://www.googleapis.com/auth/drive.file",
    ]

    creds = None

    # The file token.pickle stores the user's access and refresh tokens, and is created automatically when the authorization flow completes for the first time.
    if os.path.exists(f"{cwd}/credential/{mode}/token.pickle"):
        with open(f"{cwd}/credential/{mode}/token.pickle", "rb") as token:
            creds = pickle.load(token)

    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                f"{cwd}/credential/{mode}/{clientjson}",
                SCOPES,
            )
            creds = flow.run_local_server(port=0)

        # Save the credentials for the next run
        with open(f"{cwd}/credential/{mode}/token.pickle", "wb") as token:
            pickle.dump(creds, token)

    return build("drive", "v3", credentials=creds)


# def get_file_id():
#     service = get_gdrive_service()

#     files_prefix = "SPV_DB"
#     files_page_token = None
#     files_df = pandas.DataFrame()

#     while True:
#         files = service.files().list(q=f"mimeType = 'application/pdf' and '{folder_id}' in parents and name contains '{files_prefix}'",
#                                      pageSize=1000, pageToken=files_page_token).execute()

#         files_page_token = files.get('nextPageToken')

#         if files_df.empty:
#             files_df = pandas.DataFrame.from_dict(
#                 files.get('files'))
#         else:
#             files_df = pandas.concat([
#                 files_df, pandas.DataFrame.from_dict(files.get('files'))])

#         if not files_page_token:
#             break

#     file_id = files_df[:1]["id"].values[0]
#     return file_id

def get_file_id(fileName):
    service = get_gdrive_service()

    folder_id = '195hNR5Z9oa5-rae04xC50zRE-WRhqYgy'  # Replace 'your_folder_id' with the actual ID of the Google Drive folder

    file_name = fileName
    files_page_token = None
    files_df = pandas.DataFrame()

    while True:
        files = service.files().list(q=f"mimeType = 'application/pdf' and '{folder_id}' in parents and name contains '{file_name}'",
                                     pageSize=1000, pageToken=files_page_token).execute()

        files_page_token = files.get('nextPageToken')

        if files_df.empty:
            files_df = pandas.DataFrame.from_dict(
                files.get('files'))
        else:
            files_df = pandas.concat([
                files_df, pandas.DataFrame.from_dict(files.get('files'))])

        if not files_page_token:
            break

    if not files_df.empty:
        file_id = files_df.iloc[0]["id"]
        return file_id
    else:
        return None  # Return None if the file was not found in the folder



def download_file(file_id):
    service = get_gdrive_service()
    service.permissions().create(
        body={"role": "reader", "type": "anyone"}, fileId=file_id
    ).execute()

    def get_confirm_token(response):
        for key, value in response.cookies.items():
            if key.startswith("download_warning"):
                return value
        return None

    def save_response_content(response):
        CHUNK_SIZE = 32768

        # get the file size from Content-length response header
        file_size = int(response.headers.get("Content-Length", 0))

        # extract Content disposition from response headers
        content_disposition = response.headers.get("content-disposition")

        # parse filename
        filename = re.findall('filename="(.+)"', content_disposition)[0]
        print("[+] File size:", file_size)
        print("[+] File name:", filename)
        progress = tqdm(
            response.iter_content(CHUNK_SIZE),
            f"Downloading {filename}",
            total=file_size,
            unit="Byte",
            unit_scale=True,
            unit_divisor=1024,
        )

        

        if not os.path.exists(f"{cwd}/output/{datestring}"):
            os.makedirs(f"{cwd}/output/{datestring}")
        with open(f"{cwd}/output/{datestring}/{filename}", "wb") as f:
            for chunk in progress:
                if chunk:
                    # filter out keep-alive new chunks
                    f.write(chunk)

                    # update the progress bar
                    progress.update(len(chunk))
        progress.close()

        return filename

    # base URL for download
    URL = "https://drive.google.com/uc?"

    # init a HTTP session
    session = requests.Session()

    # make a request
    response = session.get(URL, params={"id": file_id}, stream=True)
    print("[+] Downloading", response.url)

    # get confirmation token
    token = get_confirm_token(response)
    if token:
        params = {"id": file_id, "confirm": token}
        response = session.get(URL, params=params, stream=True)

    # download to disk
    filename = save_response_content(response)

    service.permissions().delete(
        fileId=file_id, permissionId="anyoneWithLink"
    ).execute()

    return filename
