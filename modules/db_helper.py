import pandas as pd
import sqlite3
import os
import datetime
import os



class DatabaseHelper():
    def __init__(self, db_file):
        self.database = db_file
        # create a database connection
        # if os.path.exists(db_file):
        self.conn = self.create_connection(self.database)
        # else:
        #     raise Exception("Es wurde keine Laborauswertung gefunden")


    def delete_empty_rows(self):
        with self.conn:
            sql = "DELETE FROM main WHERE Datum IS NULL OR trim(Datum) = '';"
            cur = self.conn.cursor()
            cur.execute(sql)
            self.conn.commit()

    def create_connection(self, db_file):
        """ create a database connection to the SQLite database
            specified by the db_file
        :param db_file: database file
        :return: Connection object or None
        """
        self.conn = None
        try:
            self.conn = sqlite3.connect(db_file)
        except Exception as e:
            print(e)

        return self.conn

    def get_all_probes(self):
        try:    
            with self.conn:
                self.conn.row_factory = sqlite3.Row
                sql = "SELECT * FROM main;"
                cur = self.conn.cursor()
                cur.execute(sql)
                result = [dict(row) for row in cur.fetchall()]
                return result
        except Exception as ex:
            raise Exception(f"Datenbank nicht gefunden: [{ex}]")

    def get_specific_probe(self, id, material=None, date=None):
        
        try:
            # self.conn.row_factory = sqlite3.Row
            if material and date:
                sql = f"SELECT * FROM main WHERE Kennung = '{id}' AND Materialbezeichnung = '{material}' AND Datum = '{date}';"
            else:
                sql = f"SELECT * FROM main WHERE Kennung = '{id}';"
            cur = self.conn.cursor()
            cur.execute(sql)
            result = dict(cur.fetchone())
            return result
        except Exception as ex:
            raise Exception(f"Probe nicht gefunden: [{ex}]")

    def add_laborauswertung(self, data: dict):
        values = []
        keys = []

        for key, value in data.items():
            values.append(f"'{value}'")
            keys.append(f"[{key}]")

        value_str= ', '.join(value for value in values)
        key_str = ', '.join(key for key in keys)


        with self.conn:
            cur = self.conn.cursor()
            query= f'INSERT INTO main ({key_str}) VALUES ({value_str})'
            cur.execute(query)
            self.conn.commit()

    def edit_laborauswertung(self, data: dict, kennung: str, datum: str):
        kvs = []
        sql_str = ""

        for key, value in data.items():
            if key == "Datum":
                continue
            substr = f"[{key}] = '{value}'"
            kvs.append(substr)

        sql_str = ', '.join(kvs)

        with self.conn:
            cur = self.conn.cursor()
            query= f'UPDATE main SET {sql_str} WHERE "Kennung" = "{kennung}" AND "Datum" = "{datum}";'
            cur.execute(query)
            self.conn.commit()

    def excel_to_sql(self, excel_path):
        now  = datetime.datetime.now()
        today = f"{now:%Y%m%d}"

        try:
            desktop_folder = os.path.join(os.environ['USERPROFILE'], 'Desktop')
        except:
            desktop_folder = os.path.join(os.environ['USER'], 'Desktop')

        new_db_path = os.path.join(desktop_folder, f'laborauswertung_{today}.db')

        # Datenbank erstellen
        try:
            conn = sqlite3.connect(new_db_path)
            print("Database Sqlite3.db formed.")
        except:
            print("Database Sqlite3.db not formed.")
            return
        
        print("Versuche, DB zu erstellen...")

        dfs = pd.read_excel(excel_path, sheet_name="Tabelle1")
        dfs.to_sql(name='main', con=conn, index=False, if_exists='replace')
        conn.commit()

        sql = "DELETE FROM main WHERE Datum IS NULL OR trim(Datum) = '';"
        sql_alter = "ALTER TABLE main ADD strukt_bemerkung VARCHAR(500) NULL;"
        cur = conn.cursor()
        cur.execute(sql)
        conn.commit()
        cur.execute(sql_alter)
        conn.commit()
        conn.close()
            
        print(f"Datenabk erstellt :D {new_db_path}")



if __name__ == '__main__':
    d = DatabaseHelper(r"/Users/florianzasada/Desktop/laborauswertung.db")
    d.excel_to_sql(r"\\Mac\Home\Desktop\Laborauswertung.xlsx")