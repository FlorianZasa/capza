from pydrive.drive import GoogleDrive
from pydrive.auth import GoogleAuth

class DriveHelper():
    def __init__(self) -> None:
        gauth = GoogleAuth()
        gauth.LoadCredentialsFile("mycreds.txt")
        if gauth.credentials is None:
            gauth.LocalWebserverAuth()
        elif gauth.access_token_expired:
            gauth.Refresh()
        else:
            gauth.Authorize()
        gauth.SaveCredentialsFile("mycreds.txt")
        self.drive = GoogleDrive(gauth)

    def get_version_content(self, id_file):
        metadata = dict( id = id_file )
        google_file = self.drive.CreateFile( metadata = metadata )
        google_file.GetContentFile( filename = id_file )
        content_bytes = google_file.content ; # BytesIO
        string_data = content_bytes.read().decode( 'utf-8' )
        return string_data

    def get_version(self):
        file_list = self.drive.ListFile({'q': "'1C5SYekmfuyeQBOaydd5OPt-Y0pcuexFg' in parents and trashed=false", 'maxResults': 10}).GetList()
        for file1 in file_list:
            if file1['title'] == "remote_version":
                return self.get_version_content(file1['id'])
            
        else:
            return 0

        
                


if __name__ == "__main__":
    dh = DriveHelper()
    dh.get_version()
