import pandas as pd

def get_icp_data(path_to_xlsx) -> list:
    data = pd.read_excel(path_to_xlsx)
    data = data.to_dict('records')

    return data