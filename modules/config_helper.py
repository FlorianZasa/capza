import configparser
from packaging import version
import os

from modules.drive_helper import DriveHelper

INI_FILE = ""

class ConfigHelper():
    def __init__(self, path:str) -> None:
        self.path = path
        self.config = self.read_config()
        
    def read_config(self):
        config = configparser.ConfigParser()
        print(self.path)
        with open(self.path, 'r', encoding='utf-8') as f:
            config.read_file(f)
        return config

    def get_all_config(self):
        conf_dict = {}
        for section in self.config.sections():
            for conf_key in self.config[section]:
                conf_dict[conf_key] = self.config[section][conf_key]
        return conf_dict


    def get_specific_config_value(self, key):
        for section in self.config.sections():
            for conf_key in self.config[section]:
                if conf_key == key:
                    return self.config[section][conf_key]
    
    def update_specific_value(self, key, new_value):
        for section in self.config.sections():
            for conf_key in self.config[section]:
                if conf_key == key:
                    self.config[section][conf_key] = new_value
        self._write_to_config()

    def _write_to_config(self):
        # SAVE THE SETTINGS TO THE FILE
        with open(self.path,"w", encoding='utf-8') as file_object:
            self.config.write(file_object)

    def _check_update_need(self, curr, new):
        if version.parse(curr) > version.parse(new):
            return True
        else:
            return False

    def _get_new_version(self, old_version):
        d_h = DriveHelper()
        try:
            new_version = d_h.get_version()
            if self._check_update_need(new_version, old_version):
                return new_version
            else:
                return 0
        except Exception as ex:
            raise Exception(f"Keine Verbindung zum Applikationsserver: {ex}")


    def _write_new_version(self, new_version):
        if os.path.isfile(r"./remote_version"):
            with open("./remote_version", 'w') as f:
                f.write(new_version)
        else: 
            print("Versionsfile fehlt")
            return False 
        

if __name__ == "__main__":
    cf = ConfigHelper("./config.ini")
    pnp2 = cf.update_specific_value("la_path", "")
