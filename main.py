
import datetime
from operator import truediv
import pandas as pd
from PyQt5 import QtWidgets, uic, QtGui
from PyQt5.QtWidgets import QFileDialog, QTableWidgetItem
import sys
import re

from multiprocessing import Process, Event

# import helper modules
from modules.word_helper import Word_Helper

SELECTED_PROBE = 0
SELECTED_NACHWEIS = 0

ALL_DATA_PROBE = 0
ALL_DATA_NACHWEIS = 0

class Ui(QtWidgets.QMainWindow):
    def __init__(self, parent=None):
        global ALL_DATA_NACHWEIS
        super(Ui, self).__init__(parent)
        uic.loadUi(r'capza\views\main.ui', self)
        self.show()



        self.setWindowTitle("CapZa - v0.1 - Zasada") 


        self.stackedWidget.setCurrentIndex(0)
        self.status_msg_label.setText("")
        self.file = ""
        self.today_date = datetime.date.today()

        # self.end_dateedit.setDate("2023","2","5")
        # Anfangwert aus Excel

        self.disable_buttons()

        self.nav_data_btn.clicked.connect(lambda : self.stackedWidget.setCurrentIndex(1))
        self.nav_analysis_btn.clicked.connect(lambda : self.stackedWidget.setCurrentIndex(1))
        self.nav_pnp_entry_btn.clicked.connect(lambda : self.stackedWidget.setCurrentIndex(2))
        self.nav_pnp_output_btn.clicked.connect(lambda : self.stackedWidget.setCurrentIndex(3))
        self.nav_order_form_btn.clicked.connect(lambda : self.stackedWidget.setCurrentIndex(0))

        self.find_excel_btn.clicked.connect(self.select_excel)
        self.select_probe_btn.clicked.connect(self.open_specific_probe)

        self.migrate_btn.clicked.connect(self.create_document)


        self.select_probe_btn.setEnabled(False)



    def select_excel(self):
        file = QFileDialog.getOpenFileName(self, "Öffne Excel", "C://", "Excel Files (*.xlsx *.xls)")
        self.excel_path_lineedit.setText(file[0])
        self.select_probe_btn.setEnabled(True)
        self.file = file[0]
        return file[0]

    def open_probe_win(self, dataset):
        try:
            self.probe = Probe(self)
            self.probe.show()
            self.probe.init_data(dataset)
        except Exception as ex:
            self.set_status(f"Die Excel konnte nicht geladen werden: [{ex}]")

    def create_document(self):
        print(SELECTED_NACHWEIS)
        try:
            data = {
                "projekt_nr" : str(SELECTED_PROBE["Kennung \nDiese Zeile wird zum Sortieren benötigt"]),
                "bezeichnung": str(SELECTED_NACHWEIS["Material"]).split()[1],
                # "erzeuger": str(SELECTED_NACHWEIS["Erzeuger"]).split()[1],
                # #
                # "avv": str(SELECTED_NACHWEIS["AVV"]).split()[1],
                # "menge": str(SELECTED_NACHWEIS["t"]).split()[1],
                # "heute": str(self.today_date)

            }

        # self.project_nr_lineedit.setText(str(SELECTED_PROBE["Kennung \nDiese Zeile wird zum Sortieren benötigt"]))
        # self.name_lineedit.setText(str(SELECTED_NACHWEIS["Material"]).split()[1])
        # self.person_lineedit.setText(str(SELECTED_NACHWEIS["Erzeuger"]).split()[1])
        # self.location_lineedit.setText(str(SELECTED_NACHWEIS["PLZ"]).split()[1] + " " + str(SELECTED_NACHWEIS["ORT"]).split()[1])
        # self.avv_lineedit.setText(str(SELECTED_NACHWEIS["AVV"]).split()[1])
        # self.amount_lineedit.setText(str(SELECTED_NACHWEIS["t"]).split()[1])


# self.ph_lineedit.setText(str(SELECTED_PROBE["pH-Wert"]))
#         self.leitfaehigkeit_lineedit.setText(str(SELECTED_PROBE["Leitfähigkeit (mS/cm)"]))
#         self.feuchte_lineedit.setText(str(SELECTED_PROBE["Wassergehalt %"]))
#         self.chrome_vi_lineedit.setText(str(SELECTED_PROBE["Cr 205.560 (Aqueous-Axial-iFR)"]))
#         self.lipos_ts_lineedit.setText(str(SELECTED_PROBE["Lipos TS\n%"]))
#         self.lipos_os_lineedit.setText(str(SELECTED_PROBE["Lipos FS\n%"]))
#         self.gluehverlus_lineedit.setText(str(SELECTED_PROBE["GV [%]"]))
#         self.doc_lineedit.setText(str(SELECTED_PROBE["Bezogen auf das eingewogene Material DOC mg/L "]))
#         self.tds_lineedit.setText(str(SELECTED_PROBE["\nTDS\nGesamt gelöste Stoffe (mg/L)"]))
#         self.mo_lineedit.setText(str(SELECTED_PROBE[" Bezogen auf das eingewogene Material Molybdän mg/L ………"]))
#         self.se_lineedit.setText(str(SELECTED_PROBE["Se 196.090 (Aqueous-Axial-iFR)"]))
#         self.sb_lineedit.setText(str(SELECTED_PROBE["Sb 206.833 (Aqueous-Axial-iFR)"]))
#         self.fluorid_lineedit.setText(str(SELECTED_PROBE["Fluorid mg/L"]))
#         self.chlorid_lineedit.setText(str(SELECTED_PROBE["Chlorid mg/L"]))


            wh = Word_Helper()
            wh.write_to_worfd_file(data, r"items\vorlagen\Bericht Vorlage.docx", name="NEU")
        except Exception as ex:
            self.set_status("Fehler: "+ str(ex))


    def read_excel(self):
        excel_raw = pd.read_excel(self.file)
        return excel_raw

    def empty_values(self):
        self.name_lineedit.setText("")
        self.person_lineedit.setText("")
        self.location_lineedit.setText("")
        self.avv_lineedit.setText("")
        self.amount_lineedit.setText("")

        self.ph_lineedit.setText("")
        self.leitfaehigkeit_lineedit.setText("")
        self.feuchte_lineedit.setText("")
        self.chrome_vi_lineedit.setText("")
        self.lipos_ts_lineedit.setText("-")
        self.lipos_os_lineedit.setText("")
        self.gluehverlus_lineedit.setText("")
        self.doc_lineedit.setText("")
        self.tds_lineedit.setText("")
        self.mo_lineedit.setText("")
        self.se_lineedit.setText("")
        self.sb_lineedit.setText("")
        self.fluorid_lineedit.setText("")
        self.chlorid_lineedit.setText("")
        self.toc_lineedit.setText("")
        self.ec_lineedit.setText("")

    def insert_values(self):
        global SELECTED_PROBE
        global SELECTED_NACHWEIS

        self.empty_values()

        ### in Dateneingabe
        self.project_nr_lineedit.setText(str(SELECTED_PROBE["Kennung \nDiese Zeile wird zum Sortieren benötigt"]))
        try:
            self.name_lineedit.setText(str(SELECTED_NACHWEIS["Material"]).split()[1])
        except:
            self.name_lineedit.setText("-")
        try:
            self.person_lineedit.setText(str(SELECTED_NACHWEIS["Erzeuger"]).split()[1])
        except:
            self.person_lineedit.setText("-")
        try: 
            self.location_lineedit.setText(str(SELECTED_NACHWEIS["PLZ"]).split()[1] + " " + str(SELECTED_NACHWEIS["ORT"]).split()[1])
        except:
            self.location_lineedit.setText("-")
        try:
            self.avv_lineedit.setText(str(SELECTED_NACHWEIS["AVV"]).split()[1])
        except:
            self.avv_lineedit.setText("-")
        try:
            self.amount_lineedit.setText(str(SELECTED_NACHWEIS["t"]).split()[1])
        except:
            self.amount_lineedit.setText("-")

        ### in Analysewerte
        try:
            self.ph_lineedit.setText(str(SELECTED_PROBE["pH-Wert"]))
        except:
            self.ph_lineedit.setText("-")
        try:
            self.leitfaehigkeit_lineedit.setText(str(SELECTED_PROBE["Leitfähigkeit (mS/cm)"]))
        except:
            self.leitfaehigkeit_lineedit.setText("-")
        try:
            self.feuchte_lineedit.setText(str(SELECTED_PROBE["Wassergehalt %"]))
        except:
            self.feuchte_lineedit.setText("-")
        try:
            self.chrome_vi_lineedit.setText(str(SELECTED_PROBE["Cr 205.560 (Aqueous-Axial-iFR)"]))
        except:
            self.chrome_vi_lineedit.setText("-")
        try:
            self.lipos_ts_lineedit.setText(str(SELECTED_PROBE["Lipos TS\n%"]))
        except:
            self.lipos_ts_lineedit.setText("-")
        try:
            self.lipos_os_lineedit.setText(str(SELECTED_PROBE["Lipos FS\n%"]))
        except:
            self.lipos_os_lineedit.setText("-")
        try:
            self.gluehverlus_lineedit.setText(str(SELECTED_PROBE["GV [%]"]))
        except:
            self.gluehverlus_lineedit.setText("-")
        try:
            self.doc_lineedit.setText(str(SELECTED_PROBE["Bezogen auf das eingewogene Material DOC mg/L "]))
        except:
            self.doc_lineedit.setText("-")
        try:
            self.tds_lineedit.setText(str(SELECTED_PROBE["\nTDS\nGesamt gelöste Stoffe (mg/L)"]))
        except:
            self.tds_lineedit.setText("-")
        try:
            self.mo_lineedit.setText(str(SELECTED_PROBE[" Bezogen auf das eingewogene Material Molybdän mg/L ………"]))
        except:
            self.mo_lineedit.setText("-")
        try:
            self.se_lineedit.setText(str(SELECTED_PROBE["Se 196.090 (Aqueous-Axial-iFR)"]))
        except:
            self.se_lineedit.setText("-")
        try:
            self.sb_lineedit.setText(str(SELECTED_PROBE["Sb 206.833 (Aqueous-Axial-iFR)"]))
        except:
            self.sb_lineedit.setText("-")
        try:
            self.fluorid_lineedit.setText(str(SELECTED_PROBE["Fluorid mg/L"]))
        except:
            self.fluorid_lineedit.setText("-")
        try:
            self.chlorid_lineedit.setText(str(SELECTED_PROBE["Chlorid mg/L"]))
        except:
            self.chlorid_lineedit.setText("-")
        try:
            self.toc_lineedit.setText(str(SELECTED_PROBE["TOC\n%"]))
        except:
            self.toc_lineedit.setText("-")
        try:
            self.ec_lineedit.setText(str(SELECTED_PROBE["EC\n%"]))
        except:
            self.ec_lineedit.setText("-")


        if "DK" or "UTV" or "S1" in str(SELECTED_PROBE["Kennung \nDiese Zeile wird zum Sortieren benötigt"]):
            self.disable_buttons()
        else:
            self.enable_buttons()

    def disable_buttons(self):
        self.migrate_btn.setEnabled(False)
        self.aqs_btn.setEnabled(False)
    def enable_buttons(self):
        self.migrate_btn.setEnabled(True)
        self.aqs_btn.setEnabled(True)

    def open_specific_probe(self, checked):
        global ALL_DATA_PROBE

        if isinstance(ALL_DATA_PROBE, int):
            data = self.read_excel()
            self.open_probe_win(data)
            ALL_DATA_PROBE = data
        else:
            self.open_probe_win(ALL_DATA_PROBE)

        

    def set_status(self, msg):
        self.status_msg_label.setText(msg)





class Probe(QtWidgets.QMainWindow): 
    def __init__(self, parent=None):
        super(Probe, self).__init__(parent)
        uic.loadUi(r'capza\views\select_probe.ui', self)
        
        
        self.df = ""
        self.df = ""

        self.load_probe_btn.clicked.connect(self.load_probe)
        

    def init_data(self, dataset):
        self.df = dataset
        dataset.fillna('', inplace=True)
        self.tableWidget.setRowCount(dataset.shape[0])
        self.tableWidget.setColumnCount(dataset.shape[1])
        self.tableWidget.setHorizontalHeaderLabels(dataset.columns)

        # returns pandas array object
        for row in dataset.iterrows():
            values = row[1]
            for col_index, value in enumerate(values):
                if isinstance(value, (float, int)):
                    value = '{0:0,.0f}'.format(value)
                tableItem = QTableWidgetItem(str(value))
                self.tableWidget.setItem(row[0], col_index, tableItem)

        self.tableWidget.setColumnWidth(2, 300)
        

    def load_probe(self):
        global SELECTED_PROBE
        row = self.tableWidget.currentRow()
        selected_data_serie = self.df.iloc[row]
        selected_data_dict = selected_data_serie.to_dict()


        SELECTED_PROBE = selected_data_dict
        self.differentiate_probe(str(SELECTED_PROBE["Kennung \nDiese Zeile wird zum Sortieren benötigt"]))
        self.parent().insert_values()
        self.close_window()

    def differentiate_probe(self, wert):
        global ALL_DATA_NACHWEIS
        letters, numbers = wert.split()

        for index, nummer in ALL_DATA_NACHWEIS["Nachweisnr. Werk 1"].items():
            if isinstance(letters, str):
                if isinstance(numbers, str):
                    if isinstance(nummer, str):
                        if letters and numbers in nummer:
                            self.check_in_uebersicht_nachweis(nummer)
                            return

        else:
            return
            self.check_projekt_nummer(wert)
            return "Projektnummer"

    def check_in_uebersicht_nachweis(self, projektnummer):
        print("CHECK IN ÜBERSICHT")
        global SELECTED_NACHWEIS
        nachweis_data = ALL_DATA_NACHWEIS[ALL_DATA_NACHWEIS['Nachweisnr. Werk 1'] == str(projektnummer)]
        SELECTED_NACHWEIS = nachweis_data

        


    def check_projekt_nummer(self, wert):
        df_projektnumern = pd.read_excel("items\Projektnummern.xls", sheet_name='Projekte 2022')

        

            

    def close_window(self):
        self.hide()



if __name__ == "__main__":

    ALL_DATA_NACHWEIS = pd.read_excel("items\Übersicht Nachweis.xls")

    app = QtWidgets.QApplication(sys.argv)
    win = Ui()
    sys.exit(app.exec_())

