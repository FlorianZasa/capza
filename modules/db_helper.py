import sqlite3


class DatabaseHelper():
    def __init__(self):
        self.database = r"capza/laborauswertung.db"
        # create a database connection
        self.conn = self.create_connection(self.database)


    def delete_empty_row(self):
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
            return cur.fetchall()

    def get_specific_probe(self, id):
        with self.conn:
            sql = f"SELECT * FROM main WHERE Kennung = {id};"
            cur = self.conn.cursor()
            cur.execute(sql)
            return cur.fetchone()


if __name__ == '__main__':
    d = DatabaseHelper()
    d.get_all_probes()