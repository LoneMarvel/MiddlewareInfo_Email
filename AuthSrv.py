from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pickle
import os.path
import os

SCOPES = ['https://www.googleapis.com/auth/admin.directory.user', 'https://www.googleapis.com/auth/admin.directory.user.security', 
            'https://www.googleapis.com/auth/admin.directory.group', 'https://www.googleapis.com/auth/admin.directory.group.member',
            'https://apps-apis.google.com/a/feeds/groups/']

baseDir = os.path.abspath(os.path.dirname(__file__))

def CreateService():    
    creds = None
    if os.path.exists(baseDir+'/token.pickle'):
        with open(baseDir+'/token.pickle', 'rb') as token:
            creds = pickle.load(token)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(baseDir+'/creds-sdk.json', SCOPES)
            creds = flow.run_local_server(port=0)

        with open(baseDir+'/token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('admin', 'directory_v1', credentials=creds)
    return service