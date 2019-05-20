from apiclient.http import MediaFileUpload
import sys
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

DOCUMENT_ID = '1I4xzx0ePkLE7e0AT75QCjEpcqFAHCGj4CpNtuY_vmYQ'

def gdrive_service():
    """Shows basic usage of the Drive v3 API.
    Prints the names and ids of the first 10 files the user has access to.
    """
    # If modifying these scopes, delete the file token.pickle.
    SCOPES = ['https://www.googleapis.com/auth/drive','https://www.googleapis.com/auth/documents']
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server()
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    service = build('drive', 'v3', credentials=creds)
    return service

    # Call the Drive v3 API
    results = service.files().list(
        pageSize=10, fields="nextPageToken, files(id, name)").execute()
    items = results.get('files', [])

    if not items:
        print('No files found.')
    else:
        print('Files:')
        for item in items:
            print(u'{0} ({1})'.format(item['name'], item['id']))


def gdrive_folder(folder_name):
    print ("Creating folder " + folder_name + "...")
    #We have to make a request hash to tell the google API what we're giving it
    file_metadata = {
        'name': folder_name,
        'mimeType': 'application/vnd.google-apps.folder'
    }
    file = gdrive_service.files().create(body=file_metadata).execute()
    folder_id = file.get('id')
    print ('Folder ID: %s' % folder_id)
    return folder_id


def clear_gdoc(gdocId):
	"""Deletes all content from gdocId.
    """
	document = service.documents().get(documentId=DOCUMENT_ID).execute()
	endIndex = document['body']['content'][-1]['endIndex']-1
	requests = [
	    {
	        'deleteContentRange': {
	            'range': {
	                'startIndex': 1,
	                'endIndex': endIndex,
	            }

	        }

	    },
	]
	result = service.documents().batchUpdate(
	    documentId=DOCUMENT_ID, body={'requests': requests}).execute()
	return result

def build_gdoc():
    """content of gdoc
    """
    requests = [
         {
            'insertText': {
                'location': {
                    'index': 1,
                },
                'text': 'Hello world!'
            }
        },
    ]
    
    return requests

def upload_as_gdoc(word_doc):
    # remove '.docx'
    filename = word_doc[:-5]
    print ("Uploading file " + filename + "...")
    #We have to make a request hash to tell the google API what we're giving it
    file_metadata = {
        'name': filename,
        'mimeType': 'application/vnd.google-apps.document',
        'parents': [folder_id]
    }

    #Now create the media file upload object and tell it what file to upload,
    #in this case 'test.html'
    media = MediaFileUpload(word_doc,
                            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                            resumable=True)

    #Now we're doing the actual post, creating a new file of the uploaded type
    file = gdrive_service.files().create(body=file_metadata,
                                         media_body=media).execute()
    #Because verbosity is nice
    url = file.get('id')
    print ("Created file '%s' id '%s'." % (filename, url))
    return url

def main():
    """Shows basic usage of the Docs API.
    Prints the title of a sample document.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server()
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('docs', 'v1', credentials=creds)

    # Retrieve the documents contents from the Docs service.
    # document = service.documents().get(documentId=DOCUMENT_ID).execute()
    # print('The title of the document is: {}'.format(document.get('title')))

    title = 'My Document'
    body = {
        'title': title,
    }
    doc = service.documents() \
        .create(body=body).execute()
    print('Created document with title: {0}'.format(
        doc.get('title')))
    requests = build_gdoc()
    result = service.documents().batchUpdate(
        documentId=doc.get('documentId'), body={'requests': requests}).execute()
    print('Updated document with contents')
    
    requests = build_gdoc()
    result = service.documents().batchUpdate(documentId=DOCUMENT_ID, body={'requests': requests}).execute()

if __name__ == '__main__':
    main()
