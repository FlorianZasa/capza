
import datetime, time
import pandas as pd
from PyQt5 import QtWidgets, uic, QtGui
from PyQt5.QtGui import QIcon, QPixmap
from PyQt5.QtCore import *
from PyQt5.QtWidgets import QFileDialog, QTableWidgetItem, QGraphicsDropShadowEffect, QSplashScreen
import sys

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
        uic.loadUi(r'\\mac\Home\Desktop\myBots\capza-app\capza\views\main.ui', self)



        self.setWindowTitle("CapZa - Zasada - v 0.1")
        self.setWindowIcon(QIcon(r'\\mac\Home\Desktop\myBots\capza-app\capza\assets\icon_logo.png'))


        self.stackedWidget.setCurrentIndex(0)
        self.status_msg_label.setText("")
        self.file = ""
        today_date_raw = datetime.datetime.now()
        self.today_date_string = today_date_raw.strftime(r"%d.%m.%Y")

        # self.end_dateedit.setDate("2023","2","5")
        # Anfangwert aus Excel

        self.disable_buttons()

        self.nav_data_btn.clicked.connect(lambda : self.stackedWidget.setCurrentIndex(0))
        self.nav_analysis_btn.clicked.connect(lambda : self.stackedWidget.setCurrentIndex(1))
        self.nav_pnp_entry_btn.clicked.connect(lambda : self.stackedWidget.setCurrentIndex(2))
        self.nav_pnp_output_btn.clicked.connect(lambda : self.stackedWidget.setCurrentIndex(3))
        self.nav_order_form_btn.clicked.connect(lambda : self.stackedWidget.setCurrentIndex(4))

        self.init_shadow(self.data_1)
        self.init_shadow(self.data_2)
        self.init_shadow(self.data_3)

        self.init_shadow(self.analysis_f1)
        self.init_shadow(self.analysis_f2)

        self.init_shadow(self.pnp_tapped)

        self.init_shadow(self.pnp_o_frame)

        self.init_shadow(self.order_frame)

        

        # Design Default values

        self.nav_btn_frame.setStyleSheet("QPushButton:checked"
            	                        "{"
                                            "background-color: rgb(231, 201, 0)"
                                            "color: rgb(0, 0, 0);"
                                        "}"
                                        "QPushButton:hover"
            	                        "{"
                                            "background-color: rgb(231, 201, 0)"
                                            "color: rgb(0, 0, 0);"
                                        "}")

        self.find_excel_btn.clicked.connect(self.select_excel)
        self.select_probe_btn.clicked.connect(self.open_specific_probe)

        self.migrate_btn.clicked.connect(self.create_document)


        self.select_probe_btn.setEnabled(False)

    def init_shadow(self, widget):
        effect = QGraphicsDropShadowEffect()

        effect.setOffset(0, 1)

        effect.setBlurRadius(8)

        widget.setGraphicsEffect(effect)



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
        id = "x" if self.id_check.checkState() == 2 else ""
        vorpruefung  = "x" if self.precheck_check.checkState() == 2 else ""
        

        ahv = "x" if self.ahv_check.checkState() == 2 else ""
        erzeuger = "x" if self.erzeuger_check.checkState() == 2 else ""

        nh3 = self.nh3_lineedit.text()
        h2 = self.h2_lineedit.text()
        brandtest= self.brandtest_lineedit.text()
        farbe = self.color_lineedit.text()
        konsistenz = self.consistency_lineedit.text()
        geruch = self.smell_lineedit.text()
        bemerkung = self.remark_textedit.toPlainText()
        #
        toc_check = "x" if self.toc_check.checkState() == 2 else ""
        if toc_check == "x":
            toc_check_yes = "x"
            toc_check_no = ""
        else:
            toc_check_yes = ""
            toc_check_no = "x"
        icp_check = "x" if self.icp_check.checkState() == 2 else ""
        if icp_check == "x":
            icp_check_yes = "x"
            icp_check_no = ""
        else:
            icp_check_yes = ""
            icp_check_no = "x"

        rfa_check = "x" if self.rfa_check.checkState() == 2 else ""
        if rfa_check == "x":
            rfa_check_yes = "x"
            rfa_check_no = ""
        else:
            rfa_check_yes = ""
            rfa_check_no = "x"

        fremd_analysis_check = "x" if self.fremdanalysis_check.checkState() == 2 else ""
        if fremd_analysis_check == "x":
            fremd_analysis_check_yes = "x"
            fremd_analysis_check_no = ""
        else:
            fremd_analysis_check_yes = ""
            fremd_analysis_check_no = "x"

        pic_check = "x" if self.pic_check.checkState() == 2 else ""
        if pic_check == "x":
            pic_check_yes = "x"
            pic_check_no = ""
        else:
            pic_check_yes = ""
            pic_check_no = "x"

        doc_check = "x" if self.doc_check.checkState() == 2 else ""
        if doc_check == "x":
            doc_check_yes = "x"
            doc_check_no = ""
        else:
            doc_check_yes = ""
            doc_check_no = "x"

        chlorid_check = "x" if self.chlorid_check.checkState() == 2 else ""
        if chlorid_check == "x":
            chlorid_check_yes = "x"
            chlorid_check_no = ""
        else:
            chlorid_check_yes = ""
            chlorid_check_no = "x"

        pbp_check = "x" if self.pbp_check.checkState() == 2 else ""
        if pbp_check == "x":
            pbp_check_yes = "x"
            pbp_check_no = ""
        else:
            pbp_check_yes = ""
            pbp_check_no = "x"

        pnp_check = "x" if self.pnp_check.checkState() == 2 else ""
        if pnp_check == "x":
            pnp_check_yes = "x"
            pnp_check_no = ""
        else:
            pnp_check_yes = ""
            pnp_check_no = "x"
        try:
            data = {
                "projekt_nr" : str(SELECTED_PROBE["Kennung \nDiese Zeile wird zum Sortieren benötigt"]),
                "bezeichnung": str(SELECTED_NACHWEIS["Material"]).split()[1],
                "erzeuger": str(SELECTED_NACHWEIS["Erzeuger"]).split()[1],
                #
                "id": id,
                "vorpruefung": vorpruefung,
                "ahv": ahv,
                "erzeuger": erzeuger,
                "avv": str(SELECTED_NACHWEIS["AVV"]).split()[1],
                "menge": str(SELECTED_NACHWEIS["t"]).split()[1],
                "heute": str(self.today_date_string),
                "datum": str(SELECTED_PROBE["Datum"]),
                #
                "wert": str(SELECTED_PROBE["pH-Wert"]),
                "leitfaehigkeit ": str(SELECTED_PROBE["Leitfähigkeit (mS/cm)"]),
                "doc": str(SELECTED_PROBE["Bezogen auf das eingewogene Material DOC mg/L "]),
                "molybdaen": str(SELECTED_PROBE[" Bezogen auf das eingewogene Material Molybdän mg/L ………"]),
                "selen": str(SELECTED_PROBE["Se 196.090 (Aqueous-Axial-iFR)"]),
                "antimon": "WAST IST DAS?",
                "chrom": str(SELECTED_PROBE["Cr 205.560 (Aqueous-Axial-iFR)"]),
                "tds": str(SELECTED_PROBE["\nTDS\nGesamt gelöste Stoffe (mg/L)"]),
                "chlorid": str(SELECTED_PROBE["Chlorid mg/L"]),
                "fluorid": str(SELECTED_PROBE["Fluorid mg/L"]),
                "feuchte": str(SELECTED_PROBE["Wassergehalt %"]),
                "lipos_ts": str(SELECTED_PROBE["Lipos TS\n%"]),
                "lipos_os": str(SELECTED_PROBE["Lipos FS\n%"]),
                "gluehverlust": str(SELECTED_PROBE["GV [%]"]),
                "toc": "str(SELECTED_PROBE[TOC])",
                "ec": "str(SELECTED_PROBE[EC])",
                "aoc": "WAS IST DAS?",
                "nh3": nh3,
                "h2": h2,
                "brandtest": brandtest,
                #
                "farbe": farbe,
                "konsistenz": konsistenz,
                "geruch": geruch,
                "bemerkung": bemerkung,
                #
                "rfa_yes": rfa_check_yes,
                "rfa_no": rfa_check_no,
                "doc_yes": doc_check_yes,
                "doc_no": doc_check_no,
                "icp_yes": icp_check_yes,
                "icp_no": icp_check_no,
                "toc_yes": toc_check_yes,
                "toc_no": toc_check_no,
                "cl_yes": chlorid_check_yes,
                "cl_no": chlorid_check_no,
                "pic_yes": pic_check_yes,
                "pic_no": pic_check_no,
                "fremd_yes": fremd_analysis_check_yes,
                "fremd_no": fremd_analysis_check_no,
                "pnp_yes": pnp_check_yes,
                "pnp_no": pnp_check_no,
                "pbd_yes": pbp_check_yes,
                "pbd_no": pbp_check_no
            }


            wh = Word_Helper()
            file = QFileDialog.getSaveFileName(self, 'Speicherort für Prüfbericht', 'C://', filter='*.docx')
            wh.write_to_worfd_file(data, r"\\mac\Home\Desktop\myBots\capza-app\items\vorlagen\Bericht Vorlage.docx", name=file[0])
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
        self.enable_buttons()

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

        if ("DK" or "UTV" or "S1") in str(SELECTED_PROBE["Kennung \nDiese Zeile wird zum Sortieren benötigt"]):
            self.disable_buttons()

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
        uic.loadUi(r'\\mac\Home\Desktop\myBots\capza-app\capza\views\select_probe.ui', self)
        
        
        self.df = ""
        self.df = ""

        self.load_probe_btn.clicked.connect(self.load_probe)
        

    def init_data(self, dataset):
        self.df = dataset
        dataset.fillna('', inplace=True)
        show_data = dataset[['Datum', 'Material beliebige Bezeichnung', 'Kennung \nDiese Zeile wird zum Sortieren benötigt']]
        self.tableWidget.setRowCount(show_data.shape[0])
        self.tableWidget.setColumnCount(show_data.shape[1])
        self.tableWidget.setHorizontalHeaderLabels(show_data.columns)

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
        try:
            letters, numbers = wert.split()
        except Exception as ex:
            print(ex)
            return

        for index, nummer in ALL_DATA_NACHWEIS["Nachweisnr. Werk 1"].items():
            if isinstance(letters, str):
                if isinstance(numbers, str):
                    if isinstance(nummer, str):
                        if letters and numbers in nummer:
                            self.check_in_uebersicht_nachweis(nummer)
                            return

        else:
            return

    def check_in_uebersicht_nachweis(self, projektnummer):
        global SELECTED_NACHWEIS
        nachweis_data = ALL_DATA_NACHWEIS[ALL_DATA_NACHWEIS['Nachweisnr. Werk 1'] == str(projektnummer)]
        SELECTED_NACHWEIS = nachweis_data

        


    def check_projekt_nummer(self, wert):
        df_projektnumern = pd.read_excel("items\Projektnummern.xls", sheet_name='Projekte 2022')

        

            

    def close_window(self):
        self.hide()



if __name__ == "__main__":

    ALL_DATA_NACHWEIS = pd.read_excel(r"\\mac\Home\Desktop\myBots\capza-app\items\Übersicht Nachweis.xls")

    app = QtWidgets.QApplication(sys.argv)

    # Create and display the splash screen
    splash_pix = QPixmap(r'\\mac\Home\Desktop\myBots\capza-app\capza\assets\icon_logo.png')
    splash = QSplashScreen(splash_pix, Qt.WindowStaysOnTopHint)
    splash.setMask(splash_pix.mask())
    splash.show()
    app.processEvents()


    win = Ui()
    win.show()
    splash.finish(win)
    sys.exit(app.exec_())

