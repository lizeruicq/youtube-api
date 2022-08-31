import os,pickle
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
def pickle_file_name(
        api_name = 'youtube',
        api_version = 'v3'):
    return f'token_{api_name}_{api_version}.pickle'

def load_credentials(
        api_name = 'youtube',
        api_version = 'v3'):
    pickle_file = pickle_file_name(
        api_name, api_version)

    if not os.path.exists(pickle_file):
        return None

    with open(pickle_file, 'rb') as token:
        return pickle.load(token)

def save_credentials(
        cred, api_name = 'youtube',
        api_version = 'v3'):
    pickle_file = pickle_file_name(
        api_name, api_version)

    with open(pickle_file, 'wb') as token:
        pickle.dump(cred, token)

def create_service(
        client_secret_file, scopes,
        api_name = 'youtube',
        api_version = 'v3'):
    print(client_secret_file, scopes,
        api_name, api_version,
        sep = ', ')

    cred = load_credentials(api_name, api_version)

    if not cred or not cred.valid:
        if cred and cred.expired and cred.refresh_token:
            cred.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                    client_secret_file, scopes)
            cred = flow.run_console()

    save_credentials(cred, api_name, api_version)

    try:
        service = build(api_name, api_version, credentials = cred)
        print(api_name, 'service created successfully')
        return service
    except Exception as e:
        print(api_name, 'service creation failed:', e)
        return None