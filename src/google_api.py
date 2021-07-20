import io

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload


def create_service(api_key, api_name, api_version, scope):

    creds = service_account.Credentials.from_service_account_file(api_key, scopes=scope)

    try:
        service = build(api_name, api_version, credentials=creds)
        print(api_name, 'service created successfully')
        return service
    except Exception as e:
        print('Unable to connect.')
        print(e)
        return None


def get_file_id(service, file_name, mime_type=None, parent_id=None):
    """Return the ID of a Google Drive file

    :param service: A Google Drive API service object
    :param file_name: A string, the name of the file
    :param mime_type: A string, optional MIME type of file to search for
    :param parent_id: A string, optional id of a parent folder to search in

    :return file_id: A string, file ID of the first found result
    """

    file_id = None

    query = """name='{}'
               and trashed=False
               """.format(file_name)

    if parent_id:
        query += "and parents in '{}'".format(parent_id)

    if mime_type:
        query += "and mimeType in '{}'".format(mime_type)

    try:
        results = service.files().list(
            q=query,
            fields='files(name, id)').execute()

        if len(results['files']) > 1:
            print('Multiple files found, retrieving first from list')

        file_id = results['files'][0]['id']

    except Exception as e:
        print('An error occurred: {}'.format(e))

    return file_id


def download_file(service, file_id, file_name, download_path):
    """Downloads a file from Google Drive to project folder

    :param service: A Google Drive API service object
    :param file_id: A string, file ID of the file to download
    :param file_name: A string, name and extension of file to save locally
    :param download_path: A pathlib.Path() object, folder to save file in

    :return None
    """

    file_path = download_path / file_name

    if file_path.is_file():
        print('...{} exists locally, skipping...'.format(file_path))
    else:
        request = service.files().get_media(fileId=file_id)

        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fd=fh, request=request)

        done = False

        while not done:
            status, done = downloader.next_chunk()
            print('...Downloading {}... {}'.format(file_path, status.progress() * 100))

        fh.seek(0)

        with open(file_path, 'wb') as f:
            f.write(fh.read())
            f.close()
