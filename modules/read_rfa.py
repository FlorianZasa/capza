import pandas as pd

def get_rfa_data(path_to_csv) -> list:
    data = pd.read_csv(path_to_csv, on_bad_lines='skip', sep = ';')
    data = data.to_dict('records')
    result_data = []


    for dic in data:
        if pd.isna(dic["Unnamed: 0"]):
            continue
        else:
            # del dic["Unnamed: 0"]
            result_data.append(dic)

    return result_data