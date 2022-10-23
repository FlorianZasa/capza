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
        with self.conn:
            self.conn.row_factory = sqlite3.Row
            sql = "SELECT * FROM main;"
            cur = self.conn.cursor()
            cur.execute(sql)
            result = [dict(row) for row in cur.fetchall()]
            return result

    def get_specific_probe(self, id):
        with self.conn:
            self.conn.row_factory = sqlite3.Row
            sql = f"SELECT * FROM main WHERE Kennung = '{id}';"
            cur = self.conn.cursor()
            cur.execute(sql)
            result = dict(cur.fetchone())
            return result

    def excel_to_sql(self, excel_path):
        try:
            if exists("./laborauswertung.db"):
                os.remove("./laborauswertung.db")
        except:
            pass
        with self.conn:
            dfs = pd.read_excel(excel_path, sheet_name="Tabelle1")
            dfs.to_sql(name='main', con=self.conn)
            self.conn.commit()
        print("DONE")
        self.delete_empty_rows()



if __name__ == '__main__':
    d = DatabaseHelper("./laborauswertung.db")
    d.excel_to_sql("//Mac/Home/Desktop/Laborauswertung 04.08.2022.xlsx")