import pandas as pd

def format_specific_insert_value(col_name: str, all_data: dict) -> str:
    if type(all_data) == pd.DataFrame:
        all_data = all_data.to_dict(orient='records')[0]
    return str(all_data[col_name]) if all_data[col_name] else "-"

def round_if_psbl(value) -> str:
        """ Checks if a given value is a float. If so, rounds it to 3 digits. Then returns as str

        Args:
            value (float): Float Probedata value 
                e.g.:3.123, 0.128493, ...

        Returns:
            str: Value to be set in
                e.g.: '3.123', '0.128', ...
        """

        if isinstance(value, float):
            return str(round(value, 3))
        elif value == "<LoD":
            return "-"
        else:
            return str(value)

if __name__ == '__main__':
    round_if_psbl(3.1234567)