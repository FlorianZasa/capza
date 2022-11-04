import requests
from packaging import version
import git
import subprocess


# rorepo is a Repo instance pointing to the git-python repository.
# For all you know, the first argument to Repo is a path to the repository
# you want to work with
#######

### TODO

######


class VersionHelper():
    def __init__(self) -> None:
        pass

    def run(self, curr_version):
        if version.parse(curr_version) < version.parse(self.get_new_version_from_remote()):
            return True
        else:
            return False

    def get_new_version_from_remote(self):
        remote_version = subprocess.check_output(["git", "tag"]).strip().decode('utf-8').replace("v.", "").split()[0]
        return remote_version
        

        
                


if __name__ == "__main__":
    dh = VersionHelper()
    dh.run("0.1.0")
