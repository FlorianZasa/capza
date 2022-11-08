import configparser
from packaging import version
import os


INI_FILE = ""
CPAZA_CONFIG = os.path.join(os.getenv("APPDATA"), "CapZa", "capza_config.ini")


class ConfigHelper():
    def __init__(self) -> None:
        self.path = CPAZA_CONFIG
        self.config = self.read_config()
        
    def read_config(self):
        config = configparser.ConfigParser()
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


if __name__ == "__main__":
    cf = ConfigHelper("./config.ini")
    pnp2 = cf.update_specific_value("la_path", "")
