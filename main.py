
import datetime
import threading
from docx2pdf import convert
from time import time
import pandas as pd
from PyQt5 import QtWidgets, uic
from PyQt5.QtGui import QIcon, QPixmap
from PyQt5.QtCore import QThread, pyqtSignal, Qt, QDate, QTimer
from PyQt5.QtWidgets import QFileDialog, QTableWidgetItem, QGraphicsDropShadowEffect, QSplashScreen, QProgressDialog
import numpy as np
import time
import subprocess, os, platform, sys


from win10toast import ToastNotifier

import os
dirname = os.path.dirname(__file__)

# import helper modules
from modules.word_helper import Word_Helper
from _localconfig import config
from modules.db_helper import DatabaseHelper

SELECTED_PROBE = 0
SELECTED_NACHWEIS = 0

ALL_DATA_PROBE = 0
ALL_DATA_NACHWEIS = 0
ALL_DATA_PROJECT_NR = 0

NW_PATH = ""
PNR_PATH = ""


STATUS_MSG = ""

BERICHT_FILE = ""

ALIVE, PROGRESS = True, 0


class Ui(QtWidgets.QMainWindow):
    def __init__(self, parent=None):
        global ALL_DATA_NACHWEIS
        global STATUS_MSG
        super(Ui, self).__init__(parent)
        uic.loadUi(r'./views/main.ui', self)

        self.showMaximized()
        self.init_main()

    def init_main(self):
        global STATUS_MSG

        self.nw_overview_path.setText(NW_PATH)
        self.project_nr_path.setText(PNR_PATH)
        self.disable_settings_lines()

        self.setWindowTitle(f"CapZa - Zasada - { config['version'] } ")
        self.setWindowIcon(QIcon(r'./assets/icon_logo.png'))

        self.stackedWidget.setCurrentIndex(0)

        self.save_dir = ""

        self.choose_save_bericht_path.clicked.connect(lambda: self.select_folder(self.save_bericht_path, "Wähle den Standardpfad zum Speichern aus."))

        self.logo_right_lbl.setPixmap(QPixmap("./assets/l_logo.png"))
        self.second_info_lbl.hide()
        today_date_raw = datetime.datetime.now()
        self.today_date_string = today_date_raw.strftime(r"%d.%m.%Y")

        self.notifier = ToastNotifier()

        self.error_info_btn.setStyleSheet(
            "QPushButton {"
            "background-color: transparent; "
            "border: 1px solid black;"
            "}"
            "QPushButton:hover"
            "{"
                "font-weight: bold"
            "}")

        self.error_info_btn.clicked.connect(self.showError)

        self.disable_buttons()
        self._check_for_errors()

        self.status_msg_btm.hide()

        self.brandtest_combo.currentTextChanged.connect(lambda: self.analysis_brandtest_lineedit.setText(self.brandtest_combo.currentText()))
        self.nh3_lineedit_2.textChanged.connect(lambda: self.nh3_lineedit.setText(self.nh3_lineedit_2.text()))
        self.h2_lineedit_2.textChanged.connect(lambda: self.h2_lineedit.setText(self.h2_lineedit_2.text()))

        self.pnp_output_probenahmedatum.setDate(self.get_today_qdate())
        self.pnp_out_protokoll_btn.clicked.connect(self.create_pnp_out_protokoll)


        self.nav_data_btn.clicked.connect(lambda : self.display(0))
        self.nav_analysis_btn.clicked.connect(lambda : self.display(1))
        self.nav_pnp_entry_btn.clicked.connect(lambda : self.display(2))
        self.nav_pnp_output_btn.clicked.connect(lambda : self.display(3))
        self.nav_order_form_btn.clicked.connect(lambda : self.display(4))
        self.nav_settings_btn.clicked.connect(lambda : self.display(5))

        self.save_references_btn.clicked.connect(self.save_references)

        self.init_shadow(self.data_1)
        self.init_shadow(self.data_2)
        self.init_shadow(self.data_3)
        self.init_shadow(self.select_probe_btn)
        self.init_shadow(self.find_excel_btn)
        self.init_shadow(self.migrate_btn)
        self.init_shadow(self.aqs_btn)
        self.init_shadow(self.project_data_btn)
        self.init_shadow(self.pnp_out_empty)
        self.init_shadow(self.pnp_out_protokoll_btn)
        self.init_shadow(self.auftrag_empty)
        self.init_shadow(self.auftrag_letsgo)
        self.init_shadow(self.analysis_f1)
        self.init_shadow(self.analysis_f2)
        self.init_shadow(self.pnp_o_frame)
        self.init_shadow(self.order_frame)
        self.init_shadow(self.pnp_in_allg_frame)
        

        self.end_dateedit.setDate(self.get_today_qdate())
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

        self.find_excel_btn.clicked.connect(lambda: self.select_file(self.excel_path_lineedit, self.select_probe_btn, "Wähle eine Laborauswertung aus...", "Excel Files (*.xlsx *.xls *.csv)"))
        self.select_probe_btn.clicked.connect(self.read_all_probes)

        self.choose_nw_path_btn.clicked.connect(self.choose_nw_path)
        self.choose_project_nr_btn.clicked.connect(self.choose_project_nr)


        self.migrate_btn.clicked.connect(self.create_bericht_document)
        self.aqs_btn.clicked.connect(self._no_function)

        if self.nw_overview_path.text() == "" or self.project_nr_path.text()=="":
            STATUS_MSG = "Es ist keine Nachweis Excel hinterlegt. Prüfe in den Referenzeinstellungen."
            self._check_for_errors()
            self.feedback_message("error", "Es ist keine Nachweis Excel hinterlegt. Prüfe in den Referenzeinstellungen.")

    def get_today_qdate(self):
        d,m,y = self.today_date_string.split(".")
        return QDate(int(y),int(m),int(d))

    def _no_function(self):
        global STATUS_MSG
        STATUS_MSG = "Diese Funktion steht noch nicht zu verfügung."
        self.feedback_message("info", "Diese Funktion steht noch nicht zu verfügung.")
        self._check_for_errors()

    def choose_nw_path(self):
        global NW_PATH
        global STATUS_MSG
        NW_PATH = self.select_file(self.nw_overview_path, "", "Wähle die Nachweis Liste aus...", "Excel Files (*.xlsx *.xls)")
        self.load_nachweis_data()   

    def choose_project_nr(self):
        global PNR_PATH
        PNR_PATH = self.select_file(self.project_nr_path, "", "Wähle die Projektnummernliste aus...", "Excel Files (*.xlsx *.xls)")
        self.load_project_nr()    
        self._check_for_errors()

    def load_nachweis_data(self):
        global STATUS_MSG
        global ALL_DATA_NACHWEIS
        try:
            ALL_DATA_NACHWEIS = pd.read_excel(NW_PATH)
            STATUS_MSG = ""
        except Exception as ex:
            print(ex)
            self.feedback_message("error", f"Es wurde eine falsche Liste ausgewählt. Bitte wähle eine gültige 'Nachweisliste' aus. [{ex}]")
            STATUS_MSG = str(ex)
            self._check_for_errors()

    def load_project_nr(self):
        global ALL_DATA_PROJECT_NR
        global STATUS_MSG
        ALL_DATA_PROJECT_NR = pd.read_excel(PNR_PATH)
        STATUS_MSG = ""

    def showError(self):
        self.error = Error(self)
        self.error.show()

    def _check_for_errors(self):
        global STATUS_MSG
        if STATUS_MSG != "":
            self.error_info_btn.show()
        else:
            self.error_info_btn.hide()

    def init_shadow(self, widget):
        effect = QGraphicsDropShadowEffect()
        effect.setOffset(0, 1)
        effect.setBlurRadius(8)
        widget.setGraphicsEffect(effect)

    def disable_settings_lines(self):
        self.nw_overview_path.setEnabled(False)
        self.project_nr_path.setEnabled(False)

    def select_folder(self, line, title):
        dir = QFileDialog.getExistingDirectory(self, title, "C://")
        line.setText(dir)
        self.save_folder = dir

    def select_file(self, line, button, title, file_type):
        global BERICHT_FILE
        file = QFileDialog.getOpenFileName(self, title, "C://", file_type)
        line.setText(file[0])
        BERICHT_FILE = file[0]

        # activate Button
        if button:
            button.setEnabled(True)
        self._check_for_errors()
        return file[0]
        
    def save_references(self):
        global STATUS_MSG
        try:
            nw_path = self.nw_overview_path.text()
            project_nr_path = self.project_nr_path.text()
            save_path = self.save_bericht_path.text()

            with open("_loc_conf.txt", 'w', encoding='utf-8') as f:
                f.write("{'nw_path': '"+nw_path+"',")
                f.write("'project_nr_path': '"+project_nr_path+"',")
                f.write("'save_path': '"+save_path+"',")
            self.feedback_message("success", "Neue Referenzen erfolgreich gespeichert.")
            
        except Exception as ex:
            STATUS_MSG = "Das Speichern ist fehlgeschlagen: " + str(ex)
            self._check_for_errors()
            self.feedback_message("error", f"Fehler beim Speichern: [{ex}]")

    def open_probe_win(self, dataset):
        global STATUS_MSG
        try:
            self.probe = Probe(self)
            self.probe.show()
            self.probe.init_data(dataset)
        except Exception as ex:
            STATUS_MSG = f"Die Excel konnte nicht geladen werden: [{ex}]"
            self._check_for_errors()
            self.feedback_message("error", f"Die Excel konnte nicht geladen werden: [{ex}]")

    def display(self,i):
        self.hide_second_info()
        self.stackedWidget.setCurrentIndex(i)
        if i == 1:
            self.hide_second_info()
    
        if i == 5:
            self.show_second_info("Der Pfad zur 'Nachweis Übersicht' Excel ist nur temporär und wird in Zukunft durch Echtdaten aus RAMSES ersetzt.")

    def create_bericht_document(self):
        global STATUS_MSG
        id = "x" if self.id_check_2.isChecked() else ""
        vorpruefung  = "x" if self.precheck_check_2.isChecked() else ""
        
        ahv = "x" if self.ahv_check_2.isChecked() else ""
        erzeuger = "x" if self.erzeuger_check_2.isChecked() else ""

        nh3 = str(self.nh3_lineedit.text())
        h2 = str(self.h2_lineedit.text())
        brandtest= str(self.brandtest_combo.currentText())
        farbe = str(self.color_lineedit.text())
        konsistenz = str(self.consistency_lineedit.text())
        geruch = str(self.smell_lineedit.text())
        bemerkung = str(self.remark_textedit.toPlainText())

        aoc = 0
        toc = 0
        ec = 0
        if not SELECTED_PROBE["TOC\n%"] == "":
            toc = self.round_if_psbl(float(SELECTED_PROBE["TOC\n%"]))
        else:
            toc = ""

        if not SELECTED_PROBE["EC\n%"] == "":
            ec = self.round_if_psbl(float(SELECTED_PROBE["EC\n%"]))
        else:
            ec = ""
        
        if isinstance(toc, float) and isinstance(ec, float):
            aoc = self.round_if_psbl(toc-ec)
        else:
            aoc = ""

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
    
        data = {
            "projekt_nr" : str(SELECTED_PROBE["Kennung \nDiese Zeile wird zum Sortieren benötigt"]),
            "bezeichnung": str(list(SELECTED_NACHWEIS["Material"])[0]),
            "erzeuger_name": str(list(SELECTED_NACHWEIS["Erzeuger"])[0]),
            #
            "id": id,
            "vorpruefung": vorpruefung,
            "ahv": ahv,
            "erzeuger": erzeuger,
            "avv": str(list(SELECTED_NACHWEIS["AVV"])[0]),
            "menge": str(list(SELECTED_NACHWEIS["t"])[0]),
            "heute": str(self.today_date_string),
            "datum": str(SELECTED_PROBE["Datum"]),
            #
            "wert": str(SELECTED_PROBE["pH-Wert"]),
            "leitfaehigkeit ": str(SELECTED_PROBE["Leitfähigkeit (mS/cm)"]),
            "doc": self.round_if_psbl(SELECTED_PROBE["Bezogen auf das eingewogene Material DOC mg/L "]),
            "molybdaen": self.round_if_psbl(SELECTED_PROBE["Bezogen auf das eingewogene Material DOC mg/L "]),
            "selen": self.round_if_psbl(SELECTED_PROBE["Se 196.090 (Aqueous-Axial-iFR)"]),
            "antimon": self.round_if_psbl(SELECTED_PROBE["Sb 206.833 (Aqueous-Axial-iFR)"]),
            "chrom": self.round_if_psbl(SELECTED_PROBE["Cr 205.560 (Aqueous-Axial-iFR)"]),
            "tds": self.round_if_psbl(SELECTED_PROBE["\nTDS\nGesamt gelöste Stoffe (mg/L)"]),
            "chlorid": str(SELECTED_PROBE["Chlorid mg/L"]),
            "fluorid": str(SELECTED_PROBE["Fluorid mg/L"]),
            "feuchte": str(SELECTED_PROBE["Wassergehalt %"]),
            "lipos_ts": self.round_if_psbl(SELECTED_PROBE["Lipos TS\n%"]),
            "lipos_os": self.round_if_psbl(SELECTED_PROBE["Lipos FS\n%"]),
            "gluehverlust": self.round_if_psbl(SELECTED_PROBE["GV [%]"]),
            "toc": toc,
            "ec": ec,
            "aoc": aoc,
            #
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

        word_file = self.create_word(config["bericht_vorlage"], data, "Bericht")
        try:
            self.create_pdf_bericht(word_file)
        except Exception as ex:
            self.feedback_message("attention", f"Die Word Datei wurde erfolgreich erstellt. Es konnte aber keine PDF erstellt werden. [{ex}]")
            STATUS_MSG = str(ex)
            self._check_for_errors()
            
    def create_pdf_bericht(self, wordfile):
        print(wordfile)
        file = wordfile.replace(".docx", ".pdf")
        convert(wordfile, file)
    
    def create_pnp_out_protokoll(self):
        anzahl = self.amount_analysis.currentText()
        vorlage_document = self._specific_vorlage(anzahl)
        ### get all data 

        ### get Probenehmer
        if self.probenehmer_ms_pnp_out.isChecked():
            probenehmer = "M. Segieth"
        elif self.probenehmer_sg_pnp_out.isChecked():
            probenehmer = "S. Goritz"
        elif self.probenehmer_lz_pnp_out.isChecked():
            probenehmer = "L. Zasada"
        elif self.sonstige_probenehmer.isChecked():
            probenehmer = self.sonstige_probenehmer_lineedit.text()
        else:
            probenehmer = "-"
        # get anwesende Person
        ### get Probenehmer
        if self.anw_person_ms_pnp_out.isChecked():
            anwesende_personen = "M. Segieth"
        elif self.anw_person_sg_pnp_out.isChecked():
            anwesende_personen = "S. Goritz"
        elif self.anw_person_lz_pnp_out.isChecked():
            anwesende_personen = "L. Zasada"
        elif self.sonstige_anwesende_person.isChecked():
            anwesende_personen = self.sonstige_anwesende_person_lineedit.text()
        else:
            anwesende_personen = "-"


        data = {
            "datum": self.today_date_string,
            "probenehmer": probenehmer,
            "anwesende_personen": anwesende_personen,
            "output_nr": self.output_nr_lineedit.text(),
            "output_nr_1": str(int(self.output_nr_lineedit.text())+1),
            "output_nr_2": str(int(self.output_nr_lineedit.text())+2),
            "output_nr_3": str(int(self.output_nr_lineedit.text())+3),
            "output_nr_4": str(int(self.output_nr_lineedit.text())+4)
        }
        self.create_word(vorlage_document, data, "PNP Output Protokoll")

    def create_word(self, vorlage, data, dialog_file):
        global STATUS_MSG
        try:
            wh = Word_Helper()
            file = QFileDialog.getSaveFileName(self, f'Speicherort für {dialog_file}', self.save_dir, filter='*.docx')
            if file[0]:
                wh.write_to_word_file(data, vorlage, name=file[0])
                self.feedback_message("success", "Das Protokoll wurde erfolgreich erstellt.")
                return file[0]
        except Exception as ex:
            self.feedback_message("error", f"{dialog_file} konnte nicht erstellt werden. [{ex}]")
            STATUS_MSG = f"{dialog_file} konnte nicht erstellt werden: " + str(ex)
            self._check_for_errors()

    def create_aqs_document(self):
        global STATUS_MSG

        id = "x" if self.id_check.checkState() == 2 else ""
        vorpruefung  = "x" if self.precheck_check.checkState() == 2 else ""
        ahv = "x" if self.ahv_check.checkState() == 2 else ""
        erzeuger = "x" if self.erzeuger_check.checkState() == 2 else ""


        nh3 = str(self.nh3_lineedit.text())
        h2 = str(self.h2_lineedit.text())
        brandtest= str(self.brandtest_combo.currentText())
        farbe = str(self.color_lineedit.text())
        konsistenz = str(self.consistency_lineedit.text())
        geruch = str(self.smell_lineedit.text())
        bemerkung = str(self.remark_textedit.toPlainText())

        aoc = 0
        toc = 0
        ec = 0
        if not SELECTED_PROBE["TOC\n%"] == "":
            toc = self.round_if_psbl(float(SELECTED_PROBE["TOC\n%"]))
        else:
            toc = ""

        if not SELECTED_PROBE["EC\n%"] == "":
            ec = self.round_if_psbl(float(SELECTED_PROBE["EC\n%"]))
        else:
            ec = ""
        
        if isinstance(toc, float) and isinstance(ec, float):
            aoc = self.round_if_psbl(toc-ec)
        else:
            aoc = ""

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
        data = {
                "projekt_nr" : str(SELECTED_PROBE["Kennung \nDiese Zeile wird zum Sortieren benötigt"]),
                "bezeichnung": str(SELECTED_NACHWEIS["Material"]).split()[1],
                "erzeuger_name": str(SELECTED_NACHWEIS["Erzeuger"]).split()[1],
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
                "doc": self.round_if_psbl(SELECTED_PROBE["Bezogen auf das eingewogene Material DOC mg/L "]),
                "molybdaen": self.round_if_psbl(SELECTED_PROBE["Bezogen auf das eingewogene Material DOC mg/L "]),
                "selen": self.round_if_psbl(SELECTED_PROBE["Se 196.090 (Aqueous-Axial-iFR)"]),
                "antimon": self.round_if_psbl(SELECTED_PROBE["Sb 206.833 (Aqueous-Axial-iFR)"]),
                "chrom": self.round_if_psbl(SELECTED_PROBE["Cr 205.560 (Aqueous-Axial-iFR)"]),
                "tds": self.round_if_psbl(SELECTED_PROBE["\nTDS\nGesamt gelöste Stoffe (mg/L)"]),
                "chlorid": str(SELECTED_PROBE["Chlorid mg/L"]),
                "fluorid": str(SELECTED_PROBE["Fluorid mg/L"]),
                "feuchte": str(SELECTED_PROBE["Wassergehalt %"]),
                "lipos_ts": self.round_if_psbl(SELECTED_PROBE["Lipos TS\n%"]),
                "lipos_os": self.round_if_psbl(SELECTED_PROBE["Lipos FS\n%"]),
                "gluehverlust": self.round_if_psbl(SELECTED_PROBE["GV [%]"]),
                "toc": toc,
                "ec": ec,
                "aoc": aoc,
                #
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
        self.create_word("", data, "AQS")

    def file_exists_loop(self, path, limit=30):
        exists = False
        lim = 0
        while not exists or lim <= limit :
            if exists:
                return True
            else:
                os.path.exists(path)
                time.sleep(1)
            
    def round_if_psbl(self, value):
        if isinstance(value, float):
            return round(value, 3)
        else:
            return str(value)

    def read_excel(self):
        global BERICHT_FILE
        self.feedback_message("info", "Lade Probedaten...")
        try:
            excel_raw = pd.read_csv(BERICHT_FILE, on_bad_lines='skip')
        except Exception as ex:
            try:
                excel_raw = pd.read_excel(BERICHT_FILE)
            except Exception as ex:
                STATUS_MSG =f"Konnte nicht gelesen werden: {ex}"
                self.feedback_message("error", STATUS_MSG)
                self._check_for_errors()

                return

        nan_value = float("NaN")
        excel_raw.replace("", nan_value, inplace=True)
        excel_raw.dropna(how='all', axis=0, inplace=True)

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
        global STATUS_MSG

        self.empty_values()
        self.enable_buttons()

        ### in Dateneingabe
        self.project_nr_lineedit.setText(str(SELECTED_PROBE["Kennung \nDiese Zeile wird zum Sortieren benötigt"]))
        try:
            self.name_lineedit.setText(str(list(SELECTED_NACHWEIS["Material"])[0]))
        except:
            self.name_lineedit.setText("-")

        try:
            self.person_lineedit.setText(str(list(SELECTED_NACHWEIS["Erzeuger"])[0]))
        except:
            self.person_lineedit.setText("-")
        try: 
            self.location_lineedit.setText(str(list(SELECTED_NACHWEIS["PLZ"])[0]) + " " + str(list(SELECTED_NACHWEIS["ORT"])[0]))
        except:
            self.location_lineedit.setText("-")
        try:
            self.avv_lineedit.setText(str(list(SELECTED_NACHWEIS["AVV"])[0]))
        except:
            self.avv_lineedit.setText("-")
        try:
            self.amount_lineedit.setText(str(list(SELECTED_NACHWEIS["t"])[0]))
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

        date = str(SELECTED_PROBE["Datum"]).split()[0]
        date = date.split("-")
        y = date[0]
        m = date[1]
        d = date[2]
        self.probe_date.setDate(QDate(int(y), int(m), int(d)))
        self.check_start_date.setDate(QDate(int(y),int(m),int(d)))

        self.nachweisnr_lineedit.setText(str(SELECTED_PROBE["Kennung \nDiese Zeile wird zum Sortieren benötigt"]))
        STATUS_MSG = ""
        self.feedback_message("success", "Probe erfolgreich geladen.")
        self.show_second_info("Gehe zu 'Analysewerte', um die Dokumente zu erstellen. >")

    def disable_buttons(self):
        self.migrate_btn.setEnabled(False)
        self.aqs_btn.setEnabled(False)
        self.select_probe_btn.setEnabled(False)
    def enable_buttons(self):
        self.migrate_btn.setEnabled(True)
        self.aqs_btn.setEnabled(True)

    def show_second_info(self, msg):
        self.second_info_lbl.setText(msg)
        self.second_info_lbl.show()
    def hide_second_info(self):
        self.second_info_lbl.setText("")
        self.second_info_lbl.hide()

    def _get_all_probes_handler(self):
        global STATUS_MSG

        thread1 = threading.Thread(target = self.read_all_probes)
        thread1.start()
        thread2 = threading.Thread(target = self.start_progress)
        thread2.start()

    def read_all_probes(self):
        global ALL_DATA_PROBE
        try:
            if isinstance(ALL_DATA_PROBE, int):
                data = self.read_excel()
                
                data = data.loc[::-1].reset_index(drop=True)
                self.open_probe_win(data)
                ALL_DATA_PROBE = data
            else:
                self.open_probe_win(ALL_DATA_PROBE)
            ### beende QProgressWindow
            self.end_progress()
            STATUS_MSG = ""
        except Exception as ex:
            STATUS_MSG = "Es wurde entweder kein Pfad zur Excel angegeben oder ist er ist fehlerhaft. Bitte wähle zunächst die Laborauswertung aus : "+ str(ex)
            self._check_for_errors()
            self.feedback_message("error", STATUS_MSG)

    def feedback_message(self, kind, msg):
        self.status_msg_btm.setText(msg)
        if kind == "success":
            self.status_msg_btm.setStyleSheet(
                "* {"
                    "background-color: #A2E4AE;"
                    "color: #067005;"
                    "border-radius: 10px;"
                "}"
            )
        if kind == "error":
            self.status_msg_btm.setStyleSheet(
                "* {"
                    "background-color: #ffcccc;"
                    "color: #6D0808;"
                    "border-radius: 10px;"
                "}"
            )
        if kind == "info":
            self.status_msg_btm.setStyleSheet(
                "* {"
                    "background-color: #cce0ff;"
                    "color: #003380;"
                    "border-radius: 10px;"
                "}"
            )
        if kind == "attention":
            self.status_msg_btm.setStyleSheet(
                "* {"
                    "background-color: #F5DA9D;"
                    "color: #b08b35;"
                    "border-radius: 10px;"
                "}"
            )

        self.status_msg_btm.show()
        QTimer.singleShot(3000, lambda: self.status_msg_btm.hide())

    def open_file(self, path):
        if platform.system() == "Darwin":
            subprocess.call(('open', path))
        elif platform.system() == "Windows":
            os.startfile(path)
        else:
            subprocess.call(("xdg-open", path))

    def _specific_vorlage(self, anzahl):
        if anzahl == "1":
            return config["pnp_out_1"]
        elif anzahl == "2":
            return config["pnp_out_2"]
        elif anzahl == "3":
            return config["pnp_out_3"]
        elif anzahl == "4":
            return config["pnp_out_4"]
        elif anzahl == "5":
            return config["pnp_out_5"]
        else:
            return "Ungültige Angabe"
  
    def start_progress(self):
        global PROGRESS
        self.progress = QProgressDialog('Lade alle Proben. Das kann einen Moment dauern...', 'Abbrechen', 0, 20, self)
        self.progress.setWindowTitle("Lade Proben...")
        self.progress.setWindowModality(Qt.WindowModal)
        self.progress.setFixedSize(600, 200)
        self.progress.show()
        self.progress.setValue(PROGRESS)

    def end_progress(self):
        global PROGRESS
        self.progress.setValue(PROGRESS)
        self.progress.hide()



class Probe(QtWidgets.QMainWindow): 
    def __init__(self, parent=None):
        super(Probe, self).__init__(parent)
        uic.loadUi(r'./views/select_probe.ui', self)

        self.setWindowTitle(f"CapZa - Zasada - { config['version'] } - Wähle Probe")
        
        
        self.df = ""

        self.load_probe_btn.clicked.connect(self.load_probe)
        self.init_shadow(self.load_probe_btn)
        self.init_shadow(self.cancel_btn)  

    def init_data(self, dataset):
        self.df = dataset
        dataset.fillna('', inplace=True)
        show_data = dataset[['Datum', 'Material beliebige Bezeichnung', 'Kennung \nDiese Zeile wird zum Sortieren benötigt']]
        self.tableWidget.setRowCount(show_data.shape[0])
        self.tableWidget.setColumnCount(show_data.shape[1])
        self.tableWidget.setHorizontalHeaderLabels(show_data.columns)
        self.tableWidget.setColumnWidth(0, 200)

        # returns pandas array object
        for row in dataset.iterrows():
            values = row[1]
            for col_index, value in enumerate(values):
                if isinstance(value, (float, int)):
                    value = '{0:0,.0f}'.format(value)
                tableItem = QTableWidgetItem(str(value))
                self.tableWidget.setItem(row[0], col_index, tableItem)
        self.tableWidget.setColumnWidth(2, 400)

    def init_shadow(self, widget):
        effect = QGraphicsDropShadowEffect()

        effect.setOffset(0, 1)

        effect.setBlurRadius(8)

        widget.setGraphicsEffect(effect)  

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
        global ALL_DATA_NACHWEIS, STATUS_MSG
        try:
            letters, numbers = wert.split()
        except Exception as ex:
            STATUS_MSG = ex
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
        df_projektnumern = pd.read_excel(PNR_PATH, sheet_name='Projekte 2022')

    def close_window(self):
        self.hide()


class Error(QtWidgets.QDialog): 
    def __init__(self, parent=None):
        super(Error, self).__init__(parent)
        uic.loadUi(r'./views/error.ui', self)
        global STATUS_MSG
        self.setWindowTitle("CapZa - Zasada - v 0.1 - Fehlerbeschreibung")
        self.error_lbl.setText(STATUS_MSG)
        self.init_shadow(self.close_error_info_btn)
        self.init_shadow(self.error_msg_frame)
        self.close_error_info_btn.clicked.connect(self.close_window)

    def init_shadow(self, widget):
        effect = QGraphicsDropShadowEffect()
        effect.setOffset(0, 1)
        effect.setBlurRadius(8)
        widget.setGraphicsEffect(effect)

    def close_window(self):
        self.hide()

class Worker(QThread):
    finished = pyqtSignal()
    progress = pyqtSignal(int)


    def read_all_probes(self):
        global STATUS_MSG
        global ALL_DATA_PROBE
        """Long-running task."""
        print("START TEST")
        try:
            if isinstance(ALL_DATA_PROBE, int):
                print("LESE EXCEL")
                data = self.read_excel()
                data = data.loc[::-1].reset_index(drop=True)
                print("ÖFFNE PROBEWIN")
                self.open_probe_win(data)
                ALL_DATA_PROBE = data
            else:
                self.open_probe_win(ALL_DATA_PROBE)
            ### beende QProgressWindow
            STATUS_MSG = ""
            ALIVE = False
            PROGRESS = 100

        except Exception as ex:
            STATUS_MSG = "WORKER: Es wurde entweder kein Pfad zur Excel angegeben oder ist er ist fehlerhaft. Bitte wähle zunächst die Laborauswertung aus : "+ str(ex)
            # self._check_for_errors()
            print(STATUS_MSG)
        

    def open_probe_win(self, dataset):
        try:
            self.probe = Probe()
            self.probe.show()
            self.probe.init_data(dataset)
        except Exception as ex:
            STATUS_MSG = f"Die Excel konnte nicht geladen werden: [{ex}]"
            print(STATUS_MSG)
            # self._check_for_errors(STATUS_MSG)
            # self.feedback_message("error", f"Die Excel konnte nicht geladen werden: [{ex}]")

    def read_excel(self):
        global BERICHT_FILE
        # excel_raw = pd.read_excel(BERICHT_FILE)
        try:
            excel_raw = pd.read_csv(BERICHT_FILE, on_bad_lines='skip')
        except Exception as ex:
            print(ex)

            try:
                excel_raw = pd.read_excel(BERICHT_FILE)
            except Exception as ex:
                print(f"Konnte nicht gelesen werden: {ex}")
        # nan_value = float("NaN")
        # excel_raw.replace("", nan_value, inplace=True)
        # excel_raw.dropna(how='all', axis=0, inplace=True)
        return excel_raw

if __name__ == "__main__":
    d = {}
    try:
        f = open(r"./_loc_conf.txt", 'r', encoding='utf-8')
        whole_file = f.read()
        d = eval(whole_file)

    except Exception as ex:
        pass
    if d:
        NW_PATH = d["nw_path"]
        PNR_PATH = d["project_nr_path"]
    else:
        NW_PATH = config["overview_data"]
        PNR_PATH = config["project_nr_data"]
        

    try:
        ALL_DATA_NACHWEIS = pd.read_excel(NW_PATH)
    except Exception as ex:
        STATUS_MSG = "Es wurde keine Nachweisliste gefunden. Bitte prüfe in den Referenzeinstellungen. " + str(ex)


    app = QtWidgets.QApplication(sys.argv)
    # Create and display the splash screen
    splash_pix = QPixmap("./assets/icon_logo.png")
    splash = QSplashScreen(splash_pix, Qt.WindowStaysOnTopHint)
    splash.setMask(splash_pix.mask())
    splash.show()
    app.processEvents()

    win = Ui()
    if STATUS_MSG != "":
        win._check_for_errors()
        win.feedback_message("error", STATUS_MSG)
    
    splash.finish(win)
    win.show()
    sys.exit(app.exec_())