import pandas as pd
import sqlite3
from os.path import exists
import os, sys



class DatabaseHelper():
    def __init__(self, db_file):
        self.database = db_file
        # create a database connection
        self.conn = self.create_connection(self.database)


    def delete_empty_rows(self):
        with self.conn:
            sql = "DELETE FROM main WHERE Datum IS NULL OR trim(Datum) = '';"
            cur = self.conn.cursor()
            cur.execute(sql)
            self.conn.commit()
            print("Deleted empty rows")

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
            raise Exception("Datenbank nicht gefunden")

    def get_specific_probe(self, id):
        with self.conn:
            self.conn.row_factory = sqlite3.Row
            sql = f"SELECT * FROM main WHERE Kennung = '{id}';"
            cur = self.conn.cursor()
            cur.execute(sql)
            result = dict(cur.fetchone())
            return result

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

    def edit_laborauswertung(self, data: dict, kennung:str, datum:str):
        kvs = []
        sql_str = ""

        for key, value in data.items():
            substr = f"[{key}] = '{value}'"
            kvs.append(substr)

        sql_str = ', '.join(kvs)

        with self.conn:
            cur = self.conn.cursor()
            query= f'UPDATE main SET {sql_str} WHERE "Kennung" = "{[kennung]}" AND "Datum" = "{[datum]}";'
            cur.execute(query)
            self.conn.commit()
            

    def excel_to_sql(self, excel_path):
        try:
            if exists("./laborauswertung.db"):
                os.remove("./laborauswertung.db")
        except:
            pass
        with self.conn:
            dfs = pd.read_excel(excel_path, sheet_name="Tabelle1")
            dfs.to_sql(name='main', con=self.conn, index=False)
            self.conn.commit()
        self.delete_empty_rows()



if __name__ == '__main__':
    d = DatabaseHelper("./laborauswertung.db")
    d.add_laborauswertung({"a": "1", "b": "2", "c": "3"})