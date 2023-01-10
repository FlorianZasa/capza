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
        try:
            self.conn = self.create_connection(self.database)
        except sqlite3.OperationalError:
            raise Exception("Es wurde keine Laborauswertung gefunden")


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

    def get_specific_probe(self, id, material:str=None, date:str=None):
        

        # self.conn.row_factory = sqlite3.Row
        if material and date:
            sql = f"SELECT * FROM main WHERE material_kenn = '{id}' AND material_bez = '{material}' AND datum = '{date}';"
        else:
            sql = f"SELECT * FROM main WHERE material_kenn = '{id}';"

        cur = self.conn.cursor()
        cur.execute(sql)
        result = cur.fetchone()
        if result:
            return dict(result)         
        else:
            raise Exception(f"Diese Probe konnte in der Datenbank nicht gefunden werden.") 

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
        sql_add_bemerkung = """ALTER TABLE main ADD COLUMN strukt_bemerkung VARCHAR(500) NULL;"""
        sql_add_lipos_tara = """ALTER TABLE main ADD COLUMN lipos_tara VARCHAR(255) NULL;"""
        sql_add_lipos_auswaage = """ALTER TABLE main ADD COLUMN lipos_auswaage VARCHAR(255) NULL;"""
        cur = conn.cursor()
        cur.execute(sql)
        conn.commit()
        cur.execute(sql_add_bemerkung)
        conn.commit()
        cur.execute(sql_add_lipos_tara)
        conn.commit()
        cur.execute(sql_add_lipos_auswaage)
        conn.commit()
        conn.close()
            
        print(f"Datenabk erstellt :D {new_db_path}")

    def get_all_heading_names(self) -> dict:
        try:
            cursor = self.conn.execute("select * from main")
            db_headers = [description[0] for description in cursor.description]
            local_headers = ['datum', 'material_bez', 'material_kenn', 'wassergehalt', 'einwaage_fs', 'auswaage_fs', 'ts_der_probe', 'result_ts', 'result_wasserfakt', 'result_wasserfakt_getr', 'einwaage_sox_getr', 'einwaage_sox_frisch', 'auswaage_sox_vor_nach', 'result_lipos_ts', 'result_lipos_fs', 'result_lipos_aus_frisch', 'result_lipos_fs_ts', 'gv_tara', 'gv_einwaage', 'gv_auswaage', 'result_gv', 'fluorid', 'bemerkung', 'ph_wert', 'leitfaehigkeit', 'chlorid', 'cr_vi', 'result_tds_ges', 'tds_tara', 'tds_einwaage', 'tds_auswaage', 'result_salzfracht', 'eluat_einwaage_os', 'result_einwaage_ts', 'result_faktor', 'doc', 'molybdaen', 'toc', 'ec', 'rfa_probenbezeichnung', 'Pb', 'Pb Error', 'Ni', 'Ni Error', 'Sb', 'Sb Error', 'Sn', 'Sn Error', 'Cd', 'Cd Error', 'Cr', 'Cr Error', 'Cu', 'Cu Error', 'Fe', 'Fe Error', 'Ag', 'Ag Error', 'Al', 'Al Error', 'As', 'As Error', 'Au', 'Au Error', 'Ba', 'Ba Error', 'Bal', 'Bal Error', 'Bi', 'Bi Error', 'Ca', 'Ca Error', 'Cl', 'Cl Error', 'Co', 'Co Error', 'K', 'K Error', 'Mg', 'Mg Error', 'Mn', 'Mn Error', 'Mo', 'Mo Error', 'Nb', 'Nb Error', 'P', 'P Error', 'Pd', 'Pd Error', 'Rb', 'Rb Error', 'S', 'S Error', 'Se', 'Se Error', 'Si', 'Si Error', 'Sr', 'Sr Error', 'Ti', 'Ti Error', 'Tl', 'Tl Error', 'V', 'V Error', 'W', 'W Error', 'Zn', 'Zn Error', 'Zr', 'Zr Error', 'Br', 'Br Error', 'Feuchte Stetten zwichen 17-25% ab 2018', 'ICP ab 17.02.2022/nAs 189.042 ', 'Hg 194.227 (Aqueous-Axial-iFR)', 'Se 196.090 (Aqueous-Axial-iFR)', 'Mo 202.030 (Aqueous-Axial-iFR)', 'Cr 205.560 (Aqueous-Axial-iFR)', 'Sb 206.833 (Aqueous-Axial-iFR)', 'Zn 213.856 (Aqueous-Axial-iFR)', 'Pb 220.353 (Aqueous-Axial-iFR)', 'Cd 228.802 (Aqueous-Axial-iFR)', 'Ni 231.604 (Aqueous-Axial-iFR)', 'Ba 233.527 (Aqueous-Axial-iFR)', 'Fe 259.940 (Aqueous-Axial-iFR)', 'Ca 318.128 (Aqueous-Axial-iFR)', 'Cu 324.754 (Aqueous-Axial-iFR)', 'Al 394.401 (Aqueous-Axial-iFR)', 'Ar 404.442 (Aqueous-Axial-iFR)', 'strukt_bemerkung']
            
            res = {}
            for key in local_headers:
                for value in db_headers:
                    res[key] = value
                    db_headers.remove(value)
                    break
            return res
        
        except AttributeError as ex:
            raise Exception(f"Fehler: Eventuell ist die Datenbank nicht vorhanden oder sie ist leer: [{ex}]")

if __name__ == '__main__':
    d = DatabaseHelper(r"//Mac/Home/Desktop/laborauswertung_20221125.db")
    print(d.get_specific_probe("22-0018"))