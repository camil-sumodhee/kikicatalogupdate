from __future__ import print_function
import httplib2
import os

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

try:
    import argparse

    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/sheets.googleapis.com-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive.readonly'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Kiki Catalog Update'


def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'sheets.googleapis.com-python-kikicatalogupdate.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else:  # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials


def main():
    """Shows basic usage of the Sheets API.

    Creates a Sheets API service object and prints the names and majors of
    students in a sample spreadsheet:
    https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
    """
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())

    # Drive
    # https://drive.google.com/drive/u/0/folders/0ByCbODUKWs0ja3ZBSndmWFg2MW8
    imagefolderid = '0B8vdPd-4HdtuamtXZnR0Vk80OWc'
    servicedrive = discovery.build('drive', 'v3', http=http)

    # resultsdrive = servicedrive.files().list(
    #     pageSize=10, fields="nextPageToken, files(id, name)").execute()

    resultsdrive = servicedrive.files().list( q = "'" + imagefolderid + "' in parents",
        pageSize = 5, fields = "nextPageToken, files(id, name)").execute()

    items = resultsdrive.get('files', [])
    if not items:
        print('No files found.')
    else:
        print('Files:')
        for item in items:
            print('{0} ({1})'.format(item['name'], item['id']))

    # Spreadsheet
    discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
    service = discovery.build('sheets', 'v4', http=http,
                              discoveryServiceUrl=discoveryUrl)

    # spreadsheetId = '1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms'
    spreadsheetId = '1_0tf86scoO4lOa5UCDDYKDc10U1dJ7-W6jD5Nq4FPJE'
    rangeName = 'Catalog!B1'
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheetId, range=rangeName).execute()
    values = result.get('values', [])

    if not values:
        print('No data found.')
    else:
        print('Name:')
        for row in values:
            # Print columns A, which correspond to indices 0.
            print('%s' % (row[0]))

    # Google Cloud Storage ##
    # list files
    # https://storage.googleapis.com/kikicatalog/Images/3100009479a-misako-barbara-bolso-gris.jpg
    import subprocess
    import shlex
    gs_base_url = 'https://storage.googleapis.com/kikicatalog/Images'
    proc = subprocess.Popen(shlex.split('/home/camil/Google/google-cloud-sdk/bin/gsutil ls gs://kikicatalog/Images'), stdout=subprocess.PIPE)
    result = proc.communicate()[0].decode('utf-8')

    index = 0
    values = []
    for line in result.split('\n'):
        if len(line) == 0:
            continue
        imagename = line.split('/')[4]
        print(imagename)
        imageurl = gs_base_url + '/' + imagename
        values.append(["=image(\"" + imageurl + "\")", imagename])
        index += 1
        # try:
        #
        # except:
        #     continue

    range_name = 'Catalog!A2:B' + str(index + 1)
    print('Range: ', range_name)
    body = {
        'values': values
    }
    value_input_option = 'USER_ENTERED'
    res = service.spreadsheets().values().update(
        spreadsheetId=spreadsheetId, range=range_name,
        valueInputOption=value_input_option, body=body).execute()

    print('Result: ', res)


if __name__ == '__main__':
    main()
