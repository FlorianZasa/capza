
import pandas as pd



class ProjectNrData:
    def __init__(self) -> None:
        pass
        # d = ConfigHelper()
        # config_file = d.get_all_config()
        # self.excel_file = config_file["project_nr_path"]
    def get_data(self):
        return pd.read_excel("//Mac/Home/Desktop/CapZa/Projektnummern.xls", sheet_name=None)

if __name__ == "__main__":
    p = ProjectNrData()
    print(p.get_row_by_projectid("22-0001"))
         

