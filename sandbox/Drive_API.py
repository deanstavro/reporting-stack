"""
### Kouvaya ###
The purpose of this python document is to include methods that interact with the Google Drive API.
Methods include:
   1. downloading Google spreadsheets 
   2. Moving scripts 
   3. Getting spreadsheet Ids
   4. Uploading to Google Drive
"""
from __future__ import print_function
import httplib2, os, sys, re, time, csv, io
from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage
from apiclient.http import MediaIoBaseDownload
from apiclient.http import MediaFileUpload

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# at ~/.credentials/drive-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/drive'
CLIENT_SECRET_FILE = 'client_secret_drive.json'
APPLICATION_NAME = 'Drive API To Access Google Spreadsheet Names and Ids'


"""
Returns valid user credentials from storage.
If nothing has been stored, or if the stored credentials are invalid,
the OAuth2 flow is completed to obtain the new credentials.
"""
def get_credentials():
    
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'drive-python-quickstart.json')
    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

""" 
Downloads all files in a folder

1. folderId - the id of the google folder
         Example folder id: 0Byxt_eX6xmd5TDQ4NnlSSGdoT2c
2. formatDownload - the drive acceptable download types. Refer to the google 
    drive API for acceptable mimetypes.
         Example mimetype: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
         Example mimetype: text/csv
3. fileFormat - the format you want to save your downloaded files
         Example fileformat: .xls
         Example fileformat: .csv
"""
def downloadAllSheetsFromFolder(folderId, formatDownload, fileFormat):
    
    #authorize
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('drive', 'v3', http=http)
    
    # Query daily_sheets folder for all file ids
    #'0Byxt_eX6xmd5TDQ4NnlSSGdoT2c'
    query = "'"+ folderId+"' in parents"
    results = service.files().list(
        pageSize=500,q=query,fields="nextPageToken, files(id, name)").execute()
    
    #Get the file ids and names
    items = results.get('files', [])
    if not items:
        print('No files found.')
    else:
        #Loop through files
        for item in items:
            #Get the sheet Id
            file_id=item['id']
            try:
                print("Trying a drive download")
                request = service.files().export_media(fileId=file_id,
                                             mimeType=formatDownload)
                # open this in file mode, not in byte mode, and do something with it
                fh = io.FileIO(item["id"] + fileFormat, "wb")
                downloader = MediaIoBaseDownload(fh, request)
                done = False
                while done is False:
                    status, done = downloader.next_chunk()
                    print("Download %d%%." % int(status.progress() * 100))
                    time.sleep(15)
            except:
                print("Trying a file download")
                
                try:
                    request = service.files().get_media(fileId=file_id)
                    fh = io.FileIO(item["id"] + fileFormat, "wb")
                    downloader = MediaIoBaseDownload(fh, request)
                    done = False
                    while done is False:
                        status, done = downloader.next_chunk()
                        print("Download %d%%." % int(status.progress() * 100))
                        time.sleep(5)
                except:
                    print("Was not able to download files due to some error")
    
""" 
Moves all files in a drive folder to another drive folder.

1. oldFolderId - the folder you would like to move files out of
2. newFolderId - the folder you would like to move files into
"""
def moveScripts(oldFolderId, newFolderId):   

    #authorize
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('drive', 'v3', http=http)
    
    #'0Byxt_eX6xmd5TDQ4NnlSSGdoT2c'
    query = "'"+oldFolderId+"' in parents"
    results = service.files().list(
        pageSize=500,q=query,fields="nextPageToken, files(parents, id, name)").execute()
    
    items = results.get('files', [])
    if not items:
        print('No files found.')
    else:
        #'0Byxt_eX6xmd5M1l0TTVLZ3Fzb1U'
        print('Files:')
        for item in items:
            #Get all the sheet Ids and put into a list
            file_id=item['id']
            results = service.files().update(fileId=file_id,
                                    addParents=newFolderId,
                                    removeParents=item['parents'][0],
                                    fields='id, parents').execute()
            time.sleep(12)

'''
Returns a list of spreadsheet Ids in a folder

1. folderId - Id of folder you would like file ids for
'''
def getSpreadsheetIds(folderId):

    #authorize
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('drive', 'v3', http=http)
    
    #'0Byxt_eX6xmd5M1l0TTVLZ3Fzb1U'
    query = "'"+folderId+"' in parents"
    results = service.files().list(
        pageSize=500,q=query,fields="nextPageToken, files(id, name)").execute()
    
    #print(str(results))
    items = results.get('files', [])
    if not items:
        print('No files found.')
    else:
        print('Files:')
        spreadsheetIds = []
        for item in items:
            #Get all the sheet Ids and put into a list
            file_id=item['id']
            spreadsheetIds.append(file_id)


        return spreadsheetIds

'''
Uploads a file locally to the drive.

1. fileToUpload - the name of the local file you'd like to upload
2. localPath - the path of the loval file, including the file itself
3. currentType - the type of file it is
     Example: text/csv
4. pathToUpload - the parent ID you'd like to insert the file into
5. uploadType - the type of file you'd like in your drive.
     Example: application/vnd.google-apps.spreadsheet

'''
def uploadToDrive(fileToUpload, localPath, currentType, pathToUpload, uploadType):

    #authorize
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('drive', 'v3', http=http)
    
    try:
        #localPath="/Users/KefiLabs/Desktop/paad_report/master_file.csv"
        #fileToUpload="master_file.csv"
        #'application/vnd.google-apps.spreadsheet'
        #'0Byxt_eX6xmd5aEY0Vi12Um96MHc'
        #currentType = text/csv
        file_metadata = {'name' : fileToUpload,'mimeType' : uploadType , 'parents' : [pathToUpload]}
        media = MediaFileUpload(localPath,mimetype=currentType,resumable=True)
        file = service.files().create(body=file_metadata,media_body=media,fields='id').execute()
        print('File ID: ' + file.get('id'))
        time.sleep(10)
    except:
        print("could not upload to drive")
    
