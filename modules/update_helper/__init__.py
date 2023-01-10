from urllib.request import urlopen, urlretrieve
import os
from threading import Thread
from subprocess import PIPE, Popen

INSTALLERFOLDER = os.path.join(os.environ['APPDATA'],'CapZa')

class UpdateHelper:
    def __init__(self, app_root, old_version) -> None:
        self.version_file = "https://fzasada.de/downloads/capza/update.txt"
        self.installer_file_path = "https://fzasada.de/downloads/capza/capza_installer.exe"
        self.installer_file_name = "capza_installer.exe"
        self.old_version = old_version

        self.root_dir = app_root

    def get_new_version(self) -> bool:
        config = []
        # for line in urllib.requests.
        for line in urlopen(self.version_file):
            line = line.decode("utf-8").replace(r"\n", "").replace(r"\r", "")
            config.append(line.strip())
 
        return config[0]

    def is_new_version(self) -> bool:
        config = []
        # for line in urllib.requests.
        for line in urlopen(self.version_file):
            line = line.decode("utf-8").replace(r"\n", "").replace(r"\r", "")
            config.append(line.strip())
 
        if config[0] != self.old_version:
            return True
        else:
            return False

    def clean_up(self):
        if os.path.exists(os.path.join(INSTALLERFOLDER, self.installer_file_name)):
            # delete_installer file
            os.remove(os.path.join(INSTALLERFOLDER, self.installer_file_name))

    def update(self):
        # install file from server
        urlretrieve(self.installer_file_path,  os.path.join(INSTALLERFOLDER, self.installer_file_name))
        # try:
        #     self.terminate("CapZa.exe")
        # except Exception as ex:
        #     print(ex)
        # run new installer file
        if os.path.exists(os.path.join(INSTALLERFOLDER, self.installer_file_name)):


            Popen(self.start_installer(), stdout=PIPE, stderr=PIPE, universal_newlines=True)

    def start_installer(self):
        os.system(f"start {os.path.join(INSTALLERFOLDER, self.installer_file_name)}")


    def delete_installer(self):
        os.remove(os.path.join(INSTALLERFOLDER, self.installer_file_name))



if __name__ == "__main__":
    uh = UpdateHelper(r"\\mac\Home\Desktop\myBots\capza-app\capza","1.1.11")
    print(uh.is_new_version())
    if os.path.exists(os.path.join(INSTALLERFOLDER, uh.installer_file_name)):
        uh.start_installer()

