import pyodbc
import pandas as pd

import sqlanydb

class RamsesHelper:
    def __init__(self) -> None:
        self._name = "AHV_Labor"
        self._quelle = "AHV_Labor"
        self._benutzer_id = "dba"
        self._kennwort = "sql"

    # SQL SERVER = "SELECT @@SERVERNAME"

    def connect(self): 
        #TODO:  GETESTET UND Funktioniert!!! JEDER MUSS ABER DIESEN IM ODBC ADMIN drin haben1111
        connection_string = "DSN=AHVLab17"
        conn = pyodbc.connect(connection_string)

        # conn = pyodbc.connect(
        #     'DRIVER={SQL Anywhere 17};SERVER='+server+';DATABASE='+database+';ENCRYPT=yes;UID='+username+';PWD='+ password
        # )
        return conn

    def connect_local(self):
        conn = pyodbc.connect(
            r'Driver={SQL Anywhere 17};'
            r'Server=AHVLabor;'
            r'Database=ahvlabor;'
            r'Trusted_Connection=yes;'
            r'UID=dba;'
            r'PWD=sql'
        )
        return conn

    def test_nachweis_data(self, conn):
        curs = conn.cursor()
        data = curs.execute("SELECT * FROM nachweise ")
        print(curs.fetchone())
        return data
    
    def nachweis_data(self, conn):
        data = pd.read_sql_query(f"SELECT * FROM nachweise", conn)
        return data

    def btb_data(self, conn):
        data = pd.read_sql_query(f"SELECT * FROM btbdaten", conn)
        return data

    def btb_depends_on_kennung(self, conn, kennung):
        data = pd.read_sql_query(f"SELECT * FROM btbdaten WHERE nachweisnummer = '{kennung}'", conn)
        return data


    # def data_by_

if __name__ == '__main__':
    ramses = RamsesHelper()
    conn = ramses.connect()
    print(ramses.test_nachweis_data(conn))

