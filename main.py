
import datetime
from docx2pdf import convert
import pandas as pd
from PyQt5 import QtWidgets, uic, QtCore, QtGui
from PyQt5.QtGui import QIcon, QPixmap, QFont, QStandardItemModel, QStandardItem, QIntValidator
from PyQt5.QtCore import Qt, QDate, QTimer
from PyQt5.QtWidgets import QFileDialog, QTableWidgetItem, QGraphicsDropShadowEffect, QSplashScreen, QProgressDialog, QDateEdit, QHeaderView, QComboBox, QPushButton, QCommandLinkButton
import os, sys
import re

from threading import Thread
import os
dirname = os.path.dirname(__file__)

# import helper modules
from modules.word_helper import Word_Helper
from modules.config_helper import ConfigHelper
from modules.db_helper import DatabaseHelper


LOCKFILE_PROBE = QtCore.QLockFile(QtCore.QDir.tempPath() + 'capza_probe.lock')
LOCKFILE_ERROR = QtCore.QLockFile(QtCore.QDir.tempPath() + 'capza_error.lock')


CONFIG_HELPER = ConfigHelper()
DATABASE_HELPER = DatabaseHelper(CONFIG_HELPER.get_specific_config_value("db_path"))


__version__ = CONFIG_HELPER.get_specific_config_value("version")
__update__ = False

SELECTED_PROBE = 0
SELECTED_NACHWEIS = 0

ALL_DATA_PROBE = 0
ALL_DATA_NACHWEIS = 0
ALL_DATA_PROJECT_NR = 0

NW_PATH = ""
PNR_PATH = ""
LA_PATH = ""
STANDARD_SAVE_PATH = ""
DB_PATH = ""

LA_FILTER_COUNT = 0


STATUS_MSG = []

BERICHT_FILE = ""

ALIVE, PROGRESS = True, 0


class Ui(QtWidgets.QMainWindow):
    def __init__(self, parent=None):
        global ALL_DATA_NACHWEIS
        global STATUS_MSG
        super(Ui, self).__init__(parent)
        uic.loadUi(r'./views/main.ui', self)

        self.la_changed_item_lst = {}

        self.showMaximized()
        self.init_main()

    def init_main(self):
        global STATUS_MSG

        ###STANDARDEINSTELLUNGEN:
        self.word_helper = Word_Helper()
        self.setWindowTitle(f"CapZa - Zasada - { __version__ } ")
        self.setWindowIcon(QIcon(r'./assets/icon_logo.png'))
        self.stackedWidget.setCurrentIndex(0)

        self.logo_right_lbl.setPixmap(QPixmap("./assets/l_logo.png"))
        self.second_info_lbl.hide()
        today_date_raw = datetime.datetime.now()
        self.today_date_string = today_date_raw.strftime(r"%d.%m.%Y")
        self.main_version_lbl.setText(__version__)
        self.error_info_btn.clicked.connect(self.showError)
        self.status_msg_btm.hide()
        self.hide_admin_msg_btn.clicked.connect(self.hide_admin_msg)
        

    
        # NAVIGATION:
        self.nav_data_btn.clicked.connect(lambda : self.display(0))
        self.nav_analysis_btn.clicked.connect(lambda : self.display(1))
        self.nav_pnp_entry_btn.clicked.connect(lambda : self.display(2))
        self.nav_pnp_output_btn.clicked.connect(lambda : self.display(3))
        self.nav_order_form_btn.clicked.connect(lambda : self.display(4))
        self.nav_settings_btn.clicked.connect(lambda : self.display(5))
        self.nav_laborauswertung_btn.clicked.connect(lambda : self.display(6))

        self.init_shadow(self.data_1)
        self.init_shadow(self.data_1_2)  
        self.init_shadow(self.data_2)
        self.init_shadow(self.data_3)
        self.init_shadow(self.select_probe_btn)
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
        self.init_shadow(self.clear_cache_btn)   
        self.init_shadow(self.vor_ort_frame)     

        

        ### DATENEINGABE:
        self.end_dateedit.setDate(self.get_today_qdate())
        self.load_project_nr()
        self.select_probe_btn.clicked.connect(self.read_all_probes)

        self.brandtest_combo.currentTextChanged.connect(lambda: self.analysis_brandtest_lineedit.setText(self.brandtest_combo.currentText()))
        self.nh3_lineedit_2.textChanged.connect(lambda: self.nh3_lineedit.setText(self.nh3_lineedit_2.text()))
        self.h2_lineedit_2.textChanged.connect(lambda: self.h2_lineedit.setText(self.h2_lineedit_2.text()))
        self.laborauswertung_lineedit.textChanged.connect(self.filter_laborauswertung)

        self.kennung_rb.clicked.connect(lambda: self.empty_manual_search(self.pnr_combo, self.project_nr_lineedit))
        self.pnr_rb.clicked.connect(lambda: self.empty_manual_search(self.kennung_combo, self.kennung_lineedit))

        rx = QtCore.QRegExp("\d+")
        self.kennung_lineedit.setMaxLength(4)
        self.kennung_lineedit.setValidator(QtGui.QRegExpValidator(rx))
        self.project_nr_lineedit.setMaxLength(4)
        self.project_nr_lineedit.setValidator(QtGui.QRegExpValidator(rx))
        self.search_manually_btn.clicked.connect(self.search_manual)

        ### ANALYSEWERTE:
        self.migrate_btn.clicked.connect(self.create_bericht_document)
        self.aqs_btn.clicked.connect(self._no_function)


        ### PNP Output
        int_validator = QIntValidator(0, 999999999, self)
        self.output_nr_lineedit.setValidator(int_validator)

        self.pnp_output_probenahmedatum.setDate(self.get_today_qdate())
        self.pnp_out_protokoll_btn.clicked.connect(self.create_pnp_out_protokoll)

        ### AUFTRAGSFORMULAR:
        self.autrag_load_column_view()
        self.auftrag_add_auftrag_btn.clicked.connect(self.auftrag_add_auftrag)

        ### LABORAUSWERTUNG:
        self.laborauswertung_table.doubleClicked.connect(self.edit_laborauswertung)
        self.la_edit_frame_2.hide()
        self.laborauswertung_close_edit_frame_btn.clicked.connect(self.la_cancel_edit)
        self.add_laborauswertung_btn.clicked.connect(self.add_laborauswertung)
        self.init_shadow(self.laborauswertung_table)
        self.init_shadow(self.la_edit_frame_2)


        ### REFERENZEINSTELLUNGEN:
        self.nw_overview_path.setText(NW_PATH)
        self.project_nr_path.setText(PNR_PATH)
        self.save_bericht_path.setText(STANDARD_SAVE_PATH)
        self.laborauswertung_path.setText(LA_PATH)
        self.db_path.setText(DB_PATH)
        self.disable_settings_lines()

        self.choose_project_nr_btn.clicked.connect(self.choose_project_nr)
        self.choose_nw_path_btn.clicked.connect(self.choose_nw_path)
        self.choose_laborauswertung_path_btn.clicked.connect(self.choose_la)
        self.choose_db_path_btn.clicked.connect(self.choose_db)
        self.choose_save_bericht_path.clicked.connect(lambda: self.select_folder(self.save_bericht_path, "Wähle den Standardpfad zum Speichern aus."))

        self.clear_cache_btn.clicked.connect(self.clear_cache)
        self.save_references_btn.clicked.connect(self.save_references)
        


        if self.nw_overview_path.text() == "" or self.project_nr_path.text()=="":
            STATUS_MSG.append("Es sind keine Nachweise hinterlegt. Prüfe in den Referenzeinstellungen.")
            self.feedback_message("error", STATUS_MSG)


    def hide_admin_msg(self) -> None:
        """ Hides the Admin Message Frame in the Navigation Bar
        """
        self.admin_msg_frame.hide()

    def clear_cache(self) -> None:
        """ Resets all standard variables to its default
        """
        global SELECTED_PROBE,SELECTED_NACHWEIS,ALL_DATA_PROBE,ALL_DATA_NACHWEIS,ALL_DATA_PROJECT_NR,NW_PATH,PNR_PATH,STATUS_MSG,BERICHT_FILE,ALIVE, PROGRESS
        SELECTED_PROBE = 0
        SELECTED_NACHWEIS = 0
        NW_PATH = ""
        PNR_PATH = ""
        STATUS_MSG = []
        BERICHT_FILE = ""


    def get_today_qdate(self) -> QDate:
        """ Get QDate object from current date string

        Returns:
            QDate: Contains the current Date
        """

        d,m,y = self.today_date_string.split(".")
        return QDate(int(y),int(m),int(d))

    def _set_default_style(self, widget: QtWidgets, widget_art: str) -> None:
        widget.setStyleSheet("""
            %s {
            background-color: rgb(255, 255, 255);
            color: rgb(0, 0, 0);

            border: 1px solid #C7C7C7;
            border-radius: 10px;
        }


        %s:focus {
            
            background-color: rgb(255, 253, 219);
            border: 1px solid black
        }
        """ %(widget_art, widget_art))

    def _mark_error_line(self, widget: QtWidgets, widget_art: str) -> None:
        widget.setStyleSheet("""
            border: 2px solid red;
        """)

        QTimer.singleShot(3000, lambda: self._set_default_style(widget, widget_art))

    def search_manual(self) -> None:
        global SELECTED_PROBE, SELECTED_NACHWEIS
        kennung_letters = self.kennung_combo.currentText()
        project_year = self.pnr_combo.currentText()
        kennung_nr = self.kennung_lineedit.text()
        project_nr = self.project_nr_lineedit.text()

        kennung = f"{kennung_letters} {kennung_nr}"
        projectnr = f"{project_year}-{project_nr}"

        try:
            if self.kennung_rb.isChecked():
                if kennung_letters != "-":
                    print("kennung_letters", kennung)
                    SELECTED_PROBE = DATABASE_HELPER.get_specific_probe(kennung)
                    nachweis_data = ALL_DATA_NACHWEIS[ALL_DATA_NACHWEIS['Nachweisnr. Werk 1'] == self.get_full_project_ene_number(kennung)[2]]
                    SELECTED_PROBE['Kennung_letters'] = kennung_letters
                    SELECTED_PROBE['Kennung_nr'] = kennung_nr
                    SELECTED_PROBE['Project_yr'] = "-"
                    SELECTED_PROBE['Project_nr'] = "-"
                    SELECTED_NACHWEIS = nachweis_data
                    self.insert_values()
                else:
                    self._mark_error_line(self.kennung_combo, "QComboBox")
                    raise Exception("Wählen Sie die Kennungsart aus")
            elif self.pnr_rb.isChecked():
                if project_year != "-":
                    print("1", kennung)
                    SELECTED_PROBE = DATABASE_HELPER.get_specific_probe(projectnr)
                    print("2", ALL_DATA_PROJECT_NR)
                    nachweis_data = ALL_DATA_PROJECT_NR[ALL_DATA_PROJECT_NR['Projekt-Nr.'] == projectnr]
                    nachweis_data["ORT"] = nachweis_data["Ort"]
                    nachweis_data["PLZ"] = ""
                    nachweis_data["t"] = nachweis_data["Menge [t/a]"]
                    SELECTED_PROBE['Kennung_letters'] = "-"
                    SELECTED_PROBE['Kennung_nr'] = "-"
                    SELECTED_PROBE['Project_yr'] = project_year
                    SELECTED_PROBE['Project_nr'] = project_nr
                    SELECTED_NACHWEIS = nachweis_data
                    self.insert_values()
                else:
                    self._mark_error_line(self.pnr_combo, "QComboBox")
                    raise Exception("Wählen Sie das Projektjahr aus")
            else:
                raise Exception("Es wurde keine Suchart ausgewählt. Wählen Sie eine Suchart aus.")
        except Exception as ex:
            STATUS_MSG.append(ex)
            self.feedback_message("error", STATUS_MSG)
            self.empty_values()

    def empty_manual_search(self, widget_combo, widget_line):
        # empty values
        widget_combo.setCurrentText("-")
        widget_line.setText("")


    def get_full_project_ene_number(self, nummer: str) -> str:
        """ Gets the whole , VNE, ... Projectnr. from the shortform

        Args:
            nummer (str): Shortform of the ENE Nr.
                e.g.: ENE1234

        Returns:
            str: Entire ENE Nr.
                e.g.: ENE382981234
        """

        letters, numbers = nummer.split()
        for index, nummer in ALL_DATA_NACHWEIS["Nachweisnr. Werk 1"].items():
            if isinstance(letters, str):
                if isinstance(numbers, str):
                    if isinstance(nummer, str):
                        if letters and numbers in nummer:
                            return letters, numbers, nummer
        else:
            return "/", "/", "/"

        

    def _no_function(self) -> None:
        """ Mock function for features, that are not yet implemented
        """

        global STATUS_MSG
        STATUS_MSG.append("Diese Funktion steht noch nicht zu verfügung.")
        self.feedback_message("attention", STATUS_MSG)

    def choose_nw_path(self) -> None:
        """ Choose NW_PATH from Referenzeinstellungen
        """

        global NW_PATH
        NW_PATH = self.select_file(self.nw_overview_path, "", "Wähle die Nachweis Liste aus...", "Excel Files (*.xlsx *.xls)")
        self.load_nachweis_data()

    def choose_la(self) -> None:
        """ Choose LA_PATH from Referenzeinstellungen
        """

        self.select_file(self.laborauswertung_path, "", "Wähle die Laborauswertung aus...", "Excel Files (*.xlsx *.xls)")

    def choose_db(self) -> None:
        """ Choose DB_PATH from Referenzeinstellungen
        """

        global DB_PATH
        DB_PATH = self.select_file(self.db_path, "", "Wähle die Datenbank aus...", "Databse Files (*.db)")


    def choose_project_nr(self) -> None:
        """ Choose PNR_PATH from Referenzeinstellungen
        """

        global PNR_PATH
        PNR_PATH = self.select_file(self.project_nr_path, "", "Wähle die Projektnummernliste aus...", "Excel Files (*.xlsx *.xls)") 

    def load_nachweis_data(self) -> None:
        """ Loads the data from Nachweis Übersicht.xlsx to CapZa
        """

        global STATUS_MSG
        global ALL_DATA_NACHWEIS
        try:
            ALL_DATA_NACHWEIS = pd.read_excel(NW_PATH)
            STATUS_MSG = []
        except Exception as ex:
            print(ex)
            self.feedback_message("error", [f"Es wurde eine falsche Liste ausgewählt. Bitte wähle eine gültige 'Nachweisliste' aus. [{ex}]"])
            STATUS_MSG.append(str(ex))

    def load_project_nr(self) -> None:
        """ Loads data from Projekt.xlsx to CapZa
        """

        global ALL_DATA_PROJECT_NR
        global STATUS_MSG
        try:
            ALL_DATA_PROJECT_NR = pd.read_excel(PNR_PATH, sheet_name="Projekte 2022")
            STATUS_MSG = []
        except Exception as ex:
            STATUS_MSG.append(f"Projektnummern konnten nicht geladen werden: [{ex}]")
            self.feedback_message("error", STATUS_MSG)


    def showError(self) -> None:
        """ Shows the error frame
        """
        if LOCKFILE_ERROR.tryLock(100):
            self.error = Error(self)
            self.error.show()
        else:
            pass

    def _check_for_errors(self) -> None:
        """ Checks for possible errors and schows them in case there are any
        """

        global STATUS_MSG
        if len(STATUS_MSG)>0:
            self.error_info_btn.show()
        else:
            self.error_info_btn.hide()

    def init_shadow(self, widget) -> None:
        """Sets shadow to the given widget

        Args:
            widget (QWidget): QFrame, QButton, ....
        """

        effect = QGraphicsDropShadowEffect()
        effect.setOffset(0, 1)
        effect.setBlurRadius(8)
        widget.setGraphicsEffect(effect)

    def disable_settings_lines(self) -> None:
        """ Disables Line edits from Settings 
        """

        self.nw_overview_path.setEnabled(False)
        self.project_nr_path.setEnabled(False)

    def select_folder(self, line, title:str ) -> None:
        """ Select a folder

        Args:
            line (QLineEdit): Lineedit
            title (str): Text, that will be set into the lineedit
        """

        dir = QFileDialog.getExistingDirectory(self, title, "C://")
        line.setText(dir)
        self.save_folder = dir

    def select_file(self, line: QtWidgets.QLineEdit, button: QPushButton, title: str, file_type: str) -> str:
        """ Selects a file from the file browser

        Args:
            line (QtWidgets.QLineEdit): Qline Edit that will be filled
            button (QPushButton): Button, that belongs to the Lineedit
            title (str): Title of the File
            file_type (str): Type of the file, that will be searched

        Returns:
            str: Path of the selected file
        """

        global BERICHT_FILE
        file = QFileDialog.getOpenFileName(self, title, "C://", file_type)
        line.setText(file[0])
        BERICHT_FILE = file[0]

        # activate Button
        if button:
            button.setEnabled(True)
        return file[0]

    def save_references(self) -> None:
        """ Takes all the text from the Referenzsettings and saves it to the capza_config.ini
        """

        global STATUS_MSG, ALL_DATA_PROBE
        save_path = ""
        nw_path = ""
        project_nr_path = ""
        la_path = ""
        db_path = ""

        try:
            if self.nw_overview_path.text(): 
                nw_path = self.nw_overview_path.text()
                NW_PATH = nw_path
            if self.project_nr_path.text():
                project_nr_path = self.project_nr_path.text()
                PNR_PATH = project_nr_path
            if self.save_bericht_path.text():
                save_path = self.save_bericht_path.text()
                STANDARD_SAVE_PATH = save_path
            if self.laborauswertung_path.text():
                la_path = self.laborauswertung_path.text()
                ALL_DATA_PROBE = DATABASE_HELPER.excel_to_sql(la_path)
                LA_PATH = la_path
            if self.db_path.text():
                db_path = self.db_path.text()
                DB_PATH = db_path

            references = {
                "nw_path": nw_path,
                "project_nr_path": project_nr_path,
                "save_path": save_path,
                "la_path": la_path,
                "db_path": db_path
            }

            for key, value in references.items():
                CONFIG_HELPER.update_specific_value(key, value)

            self.feedback_message("success", ["Neue Referenzen erfolgreich gespeichert."])
            STATUS_MSG = []
            
        except Exception as ex:
            STATUS_MSG.append("Das Speichern ist fehlgeschlagen: " + str(ex))
            self.feedback_message("error", f"Fehler beim Speichern: [{ex}]")

    def open_probe_win(self, dataset: list[dict]) -> None:
        """ Opens the Probe window with the entire dataset

        Args:
            dataset (dict): Dataset from the database
        """

        global STATUS_MSG
        
        try:
            if LOCKFILE_PROBE.tryLock(100):
                self.probe = Probe(self)
                self.probe.show()
                self.probe.init_data(dataset)
            else:
                print("probe already running")
                pass
        except Exception as ex:
            STATUS_MSG.append(f"Es  konnten keine Daten gefunden werden. Importiere ggf. eine Laborauswertungsexcel: [{ex}]")
            self.feedback_message("error", STATUS_MSG)
        

    def display(self, i: int) -> None:
        """ Displays the frame that is being selected in the navigation

        Args:
            i (int): Index of the Button selected in the navigation bar (connected to the Stacked frame)
        """

        self.hide_second_info()
        self.status_msg_btm.hide()

        self.stackedWidget.setCurrentIndex(i)
        if i == 1:
            self.hide_second_info()
    
        if i == 5:
            self.show_second_info("Der Pfad zur 'Nachweis Übersicht' Excel ist nur temporär und wird in Zukunft durch Echtdaten aus RAMSES ersetzt.")

        if i == 6:
            thread1 = Thread(target=self.load_laborauswertung)
            thread2 = Thread(target=self.feedback_message, args=("info", ["Laborauswertung wird geladen..."]))
            thread2.start()
            thread1.start()
            

    def create_bericht_document(self) -> None:
        """ Builds and creates the Bericht file. Therefore it gathers all data from the FE.
        """

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
        if not SELECTED_PROBE["TOC %"] == None:
            toc = self.round_if_psbl(float(SELECTED_PROBE["TOC %"]))
        else:
            toc = ""

        if not SELECTED_PROBE["EC %"] == None:
            ec = self.round_if_psbl(float(SELECTED_PROBE["EC %"]))
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

        date_str = str(SELECTED_PROBE["Datum"])
        format = r"%Y-%m-%d %H:%M:%S"
        date_dt = datetime.datetime.strptime(date_str, format)
        date = datetime.datetime.strftime(date_dt, r"%d.%m.%Y")
    
        data = {
            "projekt_nr" : str(SELECTED_PROBE["Kennung"]),
            "bezeichnung": str(list(SELECTED_NACHWEIS["Material"])[0]),
            "erzeuger_name": str(list(SELECTED_NACHWEIS["Erzeuger"])[0]),
            #
            "id": id,
            "vorpruefung": vorpruefung,
            "ahv": ahv,
            "erzeuger": erzeuger,
            "avv": self.format_avv_space_after_every_second(str(list(SELECTED_NACHWEIS["AVV"])[0])),
            "menge": str(list(SELECTED_NACHWEIS["t"])[0]),
            "heute": str(self.today_date_string),
            "datum": date,
            #
            "wert": str(SELECTED_PROBE["pH-Wert"]) if SELECTED_PROBE["pH-Wert"] != None else "",
            "leitfaehigkeit ": str(SELECTED_PROBE["Leitfähigkeit (mS/cm)"])  if SELECTED_PROBE["Leitfähigkeit (mS/cm)"] != None else "",
            "doc": self.round_if_psbl(SELECTED_PROBE["Bezogen auf das eingewogene Material DOC mg/L"])  if SELECTED_PROBE["Bezogen auf das eingewogene Material DOC mg/L"] != None else "",
            "molybdaen": self.round_if_psbl(SELECTED_PROBE["Bezogen auf das eingewogene Material DOC mg/L"]) if SELECTED_PROBE["Bezogen auf das eingewogene Material DOC mg/L"] != None else "",
            "selen": self.round_if_psbl(SELECTED_PROBE["Se 196.090 (Aqueous-Axial-iFR)"]) if SELECTED_PROBE["Se 196.090 (Aqueous-Axial-iFR)"] != None else "",
            "antimon": self.round_if_psbl(SELECTED_PROBE["Sb 206.833 (Aqueous-Axial-iFR)"]) if SELECTED_PROBE["Sb 206.833 (Aqueous-Axial-iFR)"] != None else "",
            "chrom": self.round_if_psbl(SELECTED_PROBE["Cr 205.560 (Aqueous-Axial-iFR)"]) if SELECTED_PROBE["Cr 205.560 (Aqueous-Axial-iFR)"] != None else "",
            "tds": self.round_if_psbl(SELECTED_PROBE["TDS Gesamt gelöste Stoffe (mg/L)"]) if SELECTED_PROBE["TDS Gesamt gelöste Stoffe (mg/L)"] != None else "",
            "chlorid": str(SELECTED_PROBE["Chlorid mg/L"]) if SELECTED_PROBE["Chlorid mg/L"] != None else "",
            "fluorid": str(SELECTED_PROBE["Fluorid mg/L"]) if SELECTED_PROBE["Fluorid mg/L"] != None else "",
            "feuchte": str(SELECTED_PROBE["Wassergehalt %"]) if SELECTED_PROBE["Wassergehalt %"] != None else "",
            "lipos_ts": self.round_if_psbl(SELECTED_PROBE["Lipos TS %"]) if SELECTED_PROBE["Lipos TS %"] != None else "",
            "lipos_os": self.round_if_psbl(SELECTED_PROBE["Lipos FS %"]) if SELECTED_PROBE["Lipos FS %"] != None else "",
            "gluehverlust": self.round_if_psbl(SELECTED_PROBE["GV [%]"]) if SELECTED_PROBE["GV [%]"] != None else "",
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

        word_file = self.create_word(CONFIG_HELPER.get_specific_config_value("bericht_vorlage"), data, "Bericht")        
        try:
            thread1 = Thread(target=self.word_helper.open_word, args=(word_file,))
            thread2 = Thread(target=self.feedback_message, args=("info", ["Word wird geöffnet..."]))
            thread1.start() 
            thread2.start()
        except Exception as ex:
            self.feedback_message("attention", [f"Fehler beim Erstellen der Word Datei [{ex}]"])
            STATUS_MSG.append(str(ex))

    
    def autrag_load_column_view(self) -> None:
        """ Loads the Column View
        """

        self.model = QStandardItemModel()
        self.model.setHorizontalHeaderLabels(['Projekt-/Nachweisnummer(n)', 'Probenahmedatum', 'Analyseauswahl', 'ggf. spezifische Probenbezeichnung', 'Info 2', '#'])
        self.auftrag_table_view.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.auftrag_table_view.setModel(self.model)

    def auftrag_delete_auftrag(self, row: int) -> None:
        """ Deletes a row in the Column View TODO: Get always actual Index

        Args:
            row (int): Row from the pressed Button in the Column View
        """
        self.model.removeRow(row)

    def auftrag_add_auftrag(self) -> None:
        """ Adds a row to the Column View with all its widgets
        """
        row = self.model.rowCount()
        self.model.appendRow(QStandardItem())
        # Add widgets
        _date_edit = QDateEdit(self.auftrag_table_view)
        _date_edit.setFont(QFont('Leelawadee UI', 11))
        _combo1 = QComboBox(self.auftrag_table_view)
        _combo1.addItems(["Java", "C#", "Python"])
        _combo1.setFont(QFont('Leelawadee UI', 11))
        _combo2 = QComboBox(self.auftrag_table_view)
        _combo2.addItems(["Abfall"])
        _combo2.setFont(QFont('Leelawadee UI', 11))
        _delete_btn = QPushButton("X")
        _delete_btn.clicked.connect(lambda: self.auftrag_delete_auftrag(row))
        _delete_btn.setFont(QFont('Leelawadee UI', 11))
        _delete_btn.setStyleSheet("""
            color: red;
        """)

        self.auftrag_table_view.setIndexWidget(self.model.index(row, 1), _date_edit)
        self.auftrag_table_view.setIndexWidget(self.model.index(row, 2), _combo1)
        self.auftrag_table_view.setIndexWidget(self.model.index(row, 4), _combo2)
        self.auftrag_table_view.setIndexWidget(self.model.index(row, 5), _delete_btn)

    def create_pnp_out_protokoll(self) -> None:
        """ Builds and creates the PNP-Output-Protocol. Therefore gathers all the data from the FE
        """

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

    def create_word(self, vorlage: str, data: dict, dialog_file: str) -> str:
        """ Creates a Word file based on params

        Args:
            vorlage (str): Correct Vorlage to use
            data (dict): Data that is being input into the Vorlage
            dialog_file (str): Save Folder

        Returns:
            str: Path of the new created file
        """

        global STATUS_MSG
        try:
            file = QFileDialog.getSaveFileName(self, f'Speicherort für {dialog_file}', STANDARD_SAVE_PATH, filter='*.docx')
            if file[0]:
                self.word_helper.write_to_word_file(data, vorlage, name=file[0])
                self.feedback_message("success", "Das Protokoll wurde erfolgreich erstellt.")
                return file[0]
        except Exception as ex:
            STATUS_MSG.append(f"{dialog_file} konnte nicht erstellt werden: " + str(ex))
            self.feedback_message("error", STATUS_MSG)


    def create_aqs_document(self) -> None:
        """ Builds and creates the AQS Bericht. Therefore gathers all data from FE
        """

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
        if not SELECTED_PROBE["TOC %"] == "":
            toc = self.round_if_psbl(float(SELECTED_PROBE["TOC %"]))
        else:
            toc = ""

        if not SELECTED_PROBE["EC %"] == "":
            ec = self.round_if_psbl(float(SELECTED_PROBE["EC %"]))
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
                "projekt_nr" : str(SELECTED_PROBE["Kennung"]),
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
                "doc": self.round_if_psbl(SELECTED_PROBE["Bezogen auf das eingewogene Material DOC mg/L"]),
                "molybdaen": self.round_if_psbl(SELECTED_PROBE["Bezogen auf das eingewogene Material DOC mg/L"]),
                "selen": self.round_if_psbl(SELECTED_PROBE["Se 196.090 (Aqueous-Axial-iFR)"]),
                "antimon": self.round_if_psbl(SELECTED_PROBE["Sb 206.833 (Aqueous-Axial-iFR)"]),
                "chrom": self.round_if_psbl(SELECTED_PROBE["Cr 205.560 (Aqueous-Axial-iFR)"]),
                "tds": self.round_if_psbl(SELECTED_PROBE["TDS Gesamt gelöste Stoffe (mg/L)"]),
                "chlorid": str(SELECTED_PROBE["Chlorid mg/L"]),
                "fluorid": str(SELECTED_PROBE["Fluorid mg/L"]),
                "feuchte": str(SELECTED_PROBE["Wassergehalt %"]),
                "lipos_ts": self.round_if_psbl(SELECTED_PROBE["Lipos TS %"]),
                "lipos_os": self.round_if_psbl(SELECTED_PROBE["Lipos FS %"]),
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
            
    def round_if_psbl(self, value: float) -> str:
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
        else:
            return str(value)


    def empty_values(self) -> None:
        """ Empties all LineEdits in the first two Navigations
        """

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

    def insert_values(self) -> None:
        """ Inserts all value into CapZa FE based on selected Pobe
        """

        global SELECTED_PROBE
        global SELECTED_NACHWEIS
        global STATUS_MSG

        self.empty_values()
        try:
            ### in Dateneingabe
            self.kennung_combo.setCurrentText(str(SELECTED_PROBE["Kennung_letters"]) if SELECTED_PROBE["Kennung_letters"] != None else "-")
            self.pnr_combo.setCurrentText(str(SELECTED_PROBE["Project_yr"]) if SELECTED_PROBE["Project_yr"] != None else "-")

            self.kennung_lineedit.setText(str(SELECTED_PROBE["Kennung_nr"]) if SELECTED_PROBE["Kennung_nr"] != None else "-")
            self.project_nr_lineedit.setText(str(SELECTED_PROBE["Project_nr"]) if SELECTED_PROBE["Project_nr"] != None else "-")
            self.name_lineedit.setText(str(list(SELECTED_NACHWEIS["Material"])[0])) # if SELECTED_NACHWEIS["Material"] != None else "-"
            self.person_lineedit.setText(str(list(SELECTED_NACHWEIS["Erzeuger"])[0]))
            self.location_lineedit.setText(str(list(SELECTED_NACHWEIS["PLZ"])[0]) + " " + str(list(SELECTED_NACHWEIS["ORT"])[0]))
            self.avv_lineedit.setText(self.format_avv_space_after_every_second(str(list(SELECTED_NACHWEIS["AVV"])[0])))
            self.amount_lineedit.setText("{:,}".format(list(SELECTED_NACHWEIS["t"])[0]).replace(",", "."))

            ### in Analysewerte
            self.ph_lineedit.setText(str(self.round_if_psbl(float(SELECTED_PROBE["pH-Wert"]))) if SELECTED_PROBE["pH-Wert"] != None else "-")
            self.leitfaehigkeit_lineedit.setText(str(self.round_if_psbl(float(SELECTED_PROBE["Leitfähigkeit (mS/cm)"]))) if SELECTED_PROBE["Leitfähigkeit (mS/cm)"] != None else "-")
            self.feuchte_lineedit.setText(str(self.round_if_psbl(float(SELECTED_PROBE["Wassergehalt %"]))) if SELECTED_PROBE["Wassergehalt %"] != None else "-")
            self.chrome_vi_lineedit.setText(str(self.round_if_psbl(float(SELECTED_PROBE["Cr 205.560 (Aqueous-Axial-iFR)"]))) if SELECTED_PROBE["Cr 205.560 (Aqueous-Axial-iFR)"] != None else "-")
            self.lipos_ts_lineedit.setText(str(self.round_if_psbl(float(SELECTED_PROBE["Lipos TS %"]))) if SELECTED_PROBE["Lipos TS %"] != None else "-")
            self.lipos_os_lineedit.setText(str(self.round_if_psbl(float(SELECTED_PROBE["Lipos FS %"]))) if SELECTED_PROBE["Lipos FS %"] != None else "-")
            self.gluehverlus_lineedit.setText(str(self.round_if_psbl(float(SELECTED_PROBE["GV [%]"]))) if SELECTED_PROBE["GV [%]"] != None else "-")
            self.doc_lineedit.setText(str(self.round_if_psbl(float(SELECTED_PROBE["Bezogen auf das eingewogene Material DOC mg/L"]))) if SELECTED_PROBE["Bezogen auf das eingewogene Material DOC mg/L"] != None else "-")
            self.tds_lineedit.setText(str(self.round_if_psbl(float(SELECTED_PROBE["TDS Gesamt gelöste Stoffe (mg/L)"]))) if SELECTED_PROBE["TDS Gesamt gelöste Stoffe (mg/L)"] != None else "-")
            self.mo_lineedit.setText(str(self.round_if_psbl(float(SELECTED_PROBE["Bezogen auf das eingewogene Material Molybdän mg/L"]))) if SELECTED_PROBE["Bezogen auf das eingewogene Material Molybdän mg/L"] != None else "-")
            self.se_lineedit.setText(str(self.round_if_psbl(float(SELECTED_PROBE["Se 196.090 (Aqueous-Axial-iFR)"]))) if SELECTED_PROBE["Se 196.090 (Aqueous-Axial-iFR)"] != None else "-")
            self.sb_lineedit.setText(str(self.round_if_psbl(float(SELECTED_PROBE["Sb 206.833 (Aqueous-Axial-iFR)"]))) if SELECTED_PROBE["Sb 206.833 (Aqueous-Axial-iFR)"] != None else "-")
            self.fluorid_lineedit.setText(str(self.round_if_psbl(float(SELECTED_PROBE["Fluorid mg/L"]))) if SELECTED_PROBE["Fluorid mg/L"] != None else "-")
            self.chlorid_lineedit.setText(str(self.round_if_psbl(float(SELECTED_PROBE["Chlorid mg/L"]))) if SELECTED_PROBE["Chlorid mg/L"] != None else "-")
            self.toc_lineedit.setText(str(self.round_if_psbl(float(SELECTED_PROBE["TOC %"]))) if SELECTED_PROBE["TOC %"] != None else "-")
            self.ec_lineedit.setText(str(self.round_if_psbl(float(SELECTED_PROBE["EC %"]))) if SELECTED_PROBE["EC %"] != None else "-")

            date = str(SELECTED_PROBE["Datum"]).split()[0]
            date = date.split("-")
            y = date[0]
            m = date[1]
            d = date[2]
            self.probe_date.setDate(QDate(int(y), int(m), int(d)))
            self.check_start_date.setDate(QDate(int(y),int(m),int(d)))

            self.nachweisnr_lineedit.setText(str(SELECTED_PROBE["Kennung"]))

            ### in PNP Input
            self.pnp_in_erzeuger_lineedit.setText(str(list(SELECTED_NACHWEIS["Erzeuger"])[0]))
            self.pnp_in_abfallart_textedit.setPlainText(self.format_avv_space_after_every_second(str(list(SELECTED_NACHWEIS["AVV"])[0])) + ", " + str(list(SELECTED_NACHWEIS["Material"])[0]) )

            self.feedback_message("success", ["Probe erfolgreich geladen."])
            self.show_second_info("Gehe zu 'Analysewerte', um die Dokumente zu erstellen. >")
            STATUS_MSG = []
        except Exception as ex:
            STATUS_MSG.append(f"Es konnten keine Daten ermittelt werden: [{ex}]")
            self.feedback_message("error", STATUS_MSG)


    def show_second_info(self, msg: str) -> None:
        """ Shows the second Info

        Args:
            msg (str): Info message
        """

        self.second_info_lbl.setText(msg)
        self.second_info_lbl.show()

    def hide_second_info(self) -> None:
        """ Clears and hides the info message
        """
        self.second_info_lbl.setText("")
        self.second_info_lbl.hide()

    def read_all_probes(self) -> None:
        """ Reads all Probes (from db) and save it globally
        """

        global ALL_DATA_PROBE
        try:
            STATUS_MSG = []
            if ALL_DATA_PROBE == 0:
                data = DATABASE_HELPER.get_all_probes()
                self.open_probe_win(data)
                ALL_DATA_PROBE = data
            else:
                self.open_probe_win(ALL_DATA_PROBE)
        except Exception as ex:
            STATUS_MSG.append(f"Es konnten keine Daten ermittelt werden: [{str(ex)}]")
            self.feedback_message("error", STATUS_MSG)

    def format_avv_space_after_every_second(self, avv_raw: str) -> str:
        """ Formats AVV Number: After every second charackter adds a space

        Args:
            avv_raw (str): AVV Number
                e.g.: '00000000'

        Returns:
            str: Formatted AVV Number
                e.g.: '00 00 00 00'
        """

        if len(avv_raw) > 2:
            return ' '.join(avv_raw[i:i + 2] for i in range(0, len(avv_raw), 2))
        else:
            return "/"

    def feedback_message(self, kind: str, msg: list) -> None:
        """ Shows a feedback Message colored based on the kind

        Args:
            kind (str): Kind of message
                e.g.: 'success', 'error', 'info', 'attention'
            msg (list): Message for the feedback shown to the user
                e.g.: 'Error, Please try again!'
        """

        if len(msg) > 1:
            msg = "Es bestehen mehrere Fehler. Bitte überprüfe in der Fehlerbeschreibung."
        elif len(msg) == 1:
            msg = msg[0]
        else:
            msg ="/"
        self.status_msg_btm.setText(str(msg))
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

        self._check_for_errors()
        self.status_msg_btm.show()
        QTimer.singleShot(3000, lambda: self.status_msg_btm.hide())

    def _specific_vorlage(self, anzahl: str) -> str:
        """ Return the correct Template for given amount

        Args:
            anzahl (str): Amount comes from Dropdown
                e.g.: '1', '2', ...
        Returns:
            str: Correct Vorlagen Path
        """
        if anzahl == "1":
            return CONFIG_HELPER.get_specific_config_value("pnp_out_1")
        elif anzahl == "2":
            return CONFIG_HELPER.get_specific_config_value("pnp_out_2")
        elif anzahl == "3":
            return CONFIG_HELPER.get_specific_config_value("pnp_out_3")
        elif anzahl == "4":
            return CONFIG_HELPER.get_specific_config_value("pnp_out_4")
        elif anzahl == "5":
            return CONFIG_HELPER.get_specific_config_value("pnp_out_5")
        else:
            return "Ungültige Angabe"


    def load_laborauswertung(self) -> bool:
        """ Loads the whole Laborauswertung (from db) into the FE

        Raises:
            ex: When the Probes cannot be loaded

        Returns:
            bool: True when data is being loaded; False when there is an error
        """
        global ALL_DATA_PROBE

        try:
            if ALL_DATA_PROBE == 0 or ALL_DATA_PROBE==None:
                try:
                    ALL_DATA_PROBE = DATABASE_HELPER.get_all_probes()
                except Exception as ex:
                    raise ex

            df = pd.DataFrame(ALL_DATA_PROBE)
            if df.size == 0:
                return False
            df.fillna('', inplace=True)

            self.la_search_found_lbl.setText(f"{len(df.index)} Treffer gefunden.")

            self.laborauswertung_table.setRowCount(df.shape[0])
            self.laborauswertung_table.setColumnCount(df.shape[1])
            self.laborauswertung_table.setHorizontalHeaderLabels(df.columns)

            for row in df.iterrows():
                values = row[1]
                for col_index, value in enumerate(values):
                    tableItem = QTableWidgetItem(str(value))
                    self.laborauswertung_table.setItem(row[0], col_index, tableItem)
            return True
        except Exception as ex:
            STATUS_MSG.append(f"Fehler beim Laden der Probedaten: [{str(ex)}]")
            self.feedback_message("error", STATUS_MSG)
            return False

    def edit_laborauswertung(self) -> None:
        """ Loads the selected Laborauswertung row and open the frame to edit
        """
        self.la_changed_item_lst = {}
        try:
            self.laborauswertung_edit_table.itemChanged.disconnect(self.la_handle_item_changed)
            self.laborauswertung_save_edit_frame_btn.clicked.disconnect(self.la_add_save)
        except:
            pass
        self.la_edit_title_lbl.setText("Probe bearbeiten:")

        self.laborauswertung_edit_table.setRowCount(0)
        headers = []
        row = self.laborauswertung_table.currentRow()

        for col in range(self.laborauswertung_table.columnCount()):
            headers.append(self.laborauswertung_table.horizontalHeaderItem(col).text())

        for index, title in enumerate(headers):
            try:
                self.laborauswertung_edit_table.insertRow(index)
                data = self.laborauswertung_table.item(row, index).text()
                titleItem = QTableWidgetItem(str(title))
                dataItem = QTableWidgetItem(str(data))

                flags = Qt.ItemFlags()
                flags != Qt.ItemIsEnabled
                titleItem.setFlags(flags)
                
                self.laborauswertung_edit_table.setItem(index, 0, titleItem)
                self.laborauswertung_edit_table.setItem(index, 1, dataItem)
            except Exception as ex:
                STATUS_MSG = f"Es konnten nicht alle Daten geladen werden: [{ex}]"
                self.feedback_message("error", STATUS_MSG)

        try:
            kennung = self.laborauswertung_edit_table.item(3, 1).text()
            datum = self.laborauswertung_edit_table.item(1, 1).text()
        except Exception as ex:
            STATUS_MSG =f"Es konnte keine editierbare Probe ermittelt werden."
            self.feedback_message("error", STATUS_MSG)


        self.laborauswertung_save_edit_frame_btn.clicked.connect(lambda: self.la_edit_save(kennung, datum))
        self.show_element(self.la_edit_frame_2)
        self.feedback_message("info", ["Das Editieren befindet sich zur Zeit noch in Entwicklung."])
        self.laborauswertung_edit_table.itemChanged.connect(self.la_handle_item_changed) 

    def add_laborauswertung(self) -> None:
        """ Adds a new row to the Laborauswewrtung and show the add frame
        """
        try:
            self.laborauswertung_edit_table.itemChanged.disconnect(self.la_handle_item_changed)
            self.laborauswertung_save_edit_frame_btn.clicked.disconnect(self.la_edit_save)
        except:
            pass
 
        self.la_changed_item_lst = {}

        self.la_edit_title_lbl.setText("Probe hinzufügen:")
        self.laborauswertung_edit_table.setRowCount(0)
        headers = []

        for col in range(self.laborauswertung_table.columnCount()):
            headers.append(self.laborauswertung_table.horizontalHeaderItem(col).text())

        for index, title in enumerate(headers):
            rowPosition = self.laborauswertung_edit_table.rowCount()
            if title == "index":
                pass
            else:
                self.laborauswertung_edit_table.insertRow(rowPosition)
                titleItem = QTableWidgetItem(str(title))
                flags = Qt.ItemFlags()
                flags != Qt.ItemIsEnabled
                titleItem.setFlags(flags)
                self.laborauswertung_edit_table.setItem(rowPosition, 0, titleItem)

                # Datumselect
                date_edit = QDateEdit(self.laborauswertung_edit_table)
                date_edit.setFont(QFont('Leelawadee UI', 11))
                date_edit.setDate(self.get_today_qdate())
                date_edit.dateChanged.connect(self.la_handle_item_changed)
                # self.laborauswertung_edit_table.setItem(0, 1, QTableWidgetItem(date_edit))
                self.laborauswertung_edit_table.setCellWidget(0, 1, date_edit)

        # Das Datum muss von Anfang an gesetzt werden
        #self.la_changed_item_lst[self.laborauswertung_edit_table.item(0, 0).text()] = self.laborauswertung_edit_table.item(0, 1).text()
        
        self.laborauswertung_save_edit_frame_btn.clicked.connect(self.la_add_save)
        self.show_element(self.la_edit_frame_2)
        self.feedback_message("info", ["Das Laden befindet sich zur Zeit noch in Entwicklung."])
        self.laborauswertung_edit_table.itemChanged.connect(self.la_handle_item_changed)

    def show_element(self, element: QtWidgets) -> None:
        """ Shows an Element

        Args:
            element (QtWidgets): Elemtent to be shown
                e.g.: QFrame, QPushButton, ...
        """
        element.show()

    def filter_laborauswertung(self, filter_text: str) -> None:
        """ Handles the filter event for the Laborauswertung

        Args:
            filter_text (str): Entered Filter text
        """

        anzahl = self.laborauswertung_table.rowCount()
        difference = 0
        for i in range(self.laborauswertung_table.rowCount()):
            item = self.laborauswertung_table.item(i, 2)
            if not item:
                continue
            match = filter_text.lower() not in item.text().lower()
            if not match:
                difference += 1
                break
            self.laborauswertung_table.setRowHidden(i, match)

  
        self.la_search_found_lbl.setText(f"{str(anzahl-difference)} Treffer gefunden.")

    def la_handle_item_changed(self, item) -> None:
        """ Handles the Item changed and makes varoius calculations based to input data

        Args:
            item (__type__): Changed item in Laborauswertung Table
        """

        ### Datum abfangen:
        ### TEST
        if isinstance(item, QDate):
            datum = '{0}-{1}-{2} 00:00:00'.format(item.year(), item.month(), item.day())
            self.la_changed_item_lst[self.laborauswertung_edit_table.item(0, 0).text()] = datum
            return
        else:
            try:
                ### Berechnung % TS:
                if item.row() == 4 or item.row() == 5:
                    if self.laborauswertung_edit_table.item(4, 1) and self.laborauswertung_edit_table.item(5, 1):
                        tds = float(self.laborauswertung_edit_table.item(5, 1).text()) / (float(self.laborauswertung_edit_table.item(4, 1).text())/100)
                        self.la_changed_item_lst["% TS"] = str(tds)
                        self.laborauswertung_edit_table.setItem(7, 1, QTableWidgetItem(str(tds)))
                        # Wasserfaktor
                        wasserfaktor = float(self.laborauswertung_edit_table.item(4, 1).text()) / float(self.laborauswertung_edit_table.item(5, 1).text())
                        self.la_changed_item_lst["Wasser- faktor"] = str(wasserfaktor)
                        self.laborauswertung_edit_table.setItem(8, 1, QTableWidgetItem(str(wasserfaktor)))

                ### Berechnung Wasserfaktor_getrocknet:
                if item.row() == 6:
                    if self.laborauswertung_edit_table.item(6, 1):
                        wf_getrocknet = 100 / float(self.laborauswertung_edit_table.item(6, 1).text())
                        self.la_changed_item_lst["Wasserfaktor getrocknetes Material"] = str(wf_getrocknet)
                        self.laborauswertung_edit_table.setItem(9, 1, QTableWidgetItem(str(wf_getrocknet)))

                ### Berechnung Lipos TS %:
                if item.row() == 12 or item.row() == 10:
                    if self.laborauswertung_edit_table.item(12, 1) and  self.laborauswertung_edit_table.item(10, 1):
                        lipos_ts = float(self.laborauswertung_edit_table.item(12, 1).text()) / (float(self.laborauswertung_edit_table.item(10, 1).text()) / 100)
                        self.la_changed_item_lst[r"Lipos TS %"] = str(lipos_ts)
                        self.laborauswertung_edit_table.setItem(13, 1, QTableWidgetItem(str(lipos_ts)))

                ### Berechnung Lipos FS %:
                if item.row() == 13 or item.row() == 9:
                    if self.laborauswertung_edit_table.item(13, 1) and  self.laborauswertung_edit_table.item(9, 1):
                        lipos_fs = float(self.laborauswertung_edit_table.item(13, 1).text()) / (float(self.laborauswertung_edit_table.item(9, 1).text()) / 100)
                        self.la_changed_item_lst[r"Lipos FS %"] = str(lipos_fs)
                        self.laborauswertung_edit_table.setItem(14, 1, QTableWidgetItem(str(lipos_fs)))

                ### Berechnung Lipos aus Frischsubstanz:
                if item.row() == 12 or item.row() == 11:
                    if self.laborauswertung_edit_table.item(12, 1) and  self.laborauswertung_edit_table.item(11, 1):
                        lipos_frisch = float(self.laborauswertung_edit_table.item(12, 1).text()) / (float(self.laborauswertung_edit_table.item(11, 1).text()) / 100)
                        self.la_changed_item_lst[r"% Lipos  ermittelts aus Frischsubstanz"] = str(lipos_frisch)
                        self.laborauswertung_edit_table.setItem(15, 1, QTableWidgetItem(str(lipos_frisch)))
                    
                ### Berechnung Lipos aus TS:
                if item.row() == 15 or item.row() == 8:
                    if self.laborauswertung_edit_table.item(15, 1) and  self.laborauswertung_edit_table.item(8, 1):
                        lipos_fs_ts = float(self.laborauswertung_edit_table.item(15, 1).text()) * float(self.laborauswertung_edit_table.item(8, 1).text())
                        self.la_changed_item_lst[r"% Lipos aus FS, Umrechnung auf TS"] = str(lipos_fs_ts)
                        self.laborauswertung_edit_table.setItem(16, 1, QTableWidgetItem(str(lipos_fs_ts)))

                ### Berechnung GV %:
                if item.row() == 19 or item.row() == 17 or item.row() == 18:
                    if self.laborauswertung_edit_table.item(19, 1) and self.laborauswertung_edit_table.item(17, 1) and self.laborauswertung_edit_table.item(18, 1):
                        gv = (100 - float(self.laborauswertung_edit_table.item(19, 1).text()) - float(self.laborauswertung_edit_table.item(17, 1).text())) / (float(self.laborauswertung_edit_table.item(18, 1).text()) * 100)
                        self.la_changed_item_lst[r"GV [%]"] = str(gv)
                        self.laborauswertung_edit_table.setItem(20, 1, QTableWidgetItem(str(gv)))

                ### Berechnung TDS:
                if item.row() == 30 or item.row() == 28 or item.row() == 29:
                    if self.laborauswertung_edit_table.item(30, 1) and self.laborauswertung_edit_table.item(28, 1) and self.laborauswertung_edit_table.item(29, 1):
                        gv = (float(self.laborauswertung_edit_table.item(30, 1).text()) - float(self.laborauswertung_edit_table.item(28, 1).text())) * (1000 / float(self.laborauswertung_edit_table.item(29, 1).text()) * 1000)
                        self.la_changed_item_lst[r"TDS Gesamt gelöste Stoffe (mg/L)"] = str(gv)
                        self.laborauswertung_edit_table.setItem(27, 1, QTableWidgetItem(str(gv)))


                ### Berechnung Einwaage TS
                if item.row() == 32 or item.row() == 7:
                    if self.laborauswertung_edit_table.item(32, 1) and  self.laborauswertung_edit_table.item(7, 1):
                        einwaage_ts = float(self.laborauswertung_edit_table.item(32, 1).text()) * (float(self.laborauswertung_edit_table.item(7, 1).text()) / 100)
                        self.la_changed_item_lst[r"Einwaage TS"] = str(einwaage_ts)
                        self.laborauswertung_edit_table.setItem(33, 1, QTableWidgetItem(str(einwaage_ts)))


                ### Berechnug Faktor
                if item.row() == 33 or item.row() == 32:
                    if self.laborauswertung_edit_table.item(33, 1) and  self.laborauswertung_edit_table.item(32, 1):
                        faktor = float(self.laborauswertung_edit_table.item(33, 1).text()) / float(self.laborauswertung_edit_table.item(32, 1).text())
                        self.la_changed_item_lst[r"Faktor"] = str(faktor)
                        self.laborauswertung_edit_table.setItem(34, 1, QTableWidgetItem(str(faktor)))
                

            
            except Exception as ex:
                STATUS_MSG.append(f"Eine Berechnung konnte nicht durchgeführt werden: [{ex}]")
                self.feedback_message("error", STATUS_MSG)


            self.la_changed_item_lst[self.laborauswertung_edit_table.item(item.row(), 0).text()] = item.text()

    def la_add_save(self) -> None:
        """ Adds the new Laborauswertung entry to the database
        """

        global STATUS_MSG, ALL_DATA_PROBE
        try:
            DATABASE_HELPER.add_laborauswertung(self.la_changed_item_lst)
            STATUS_MSG = []
            self.feedback_message("success", ["Erfolgreich gespeichert"])
            # TODO: Aktualisiere die Tabelle, und ALL DATA PROBE
            ALL_DATA_PROBE = DATABASE_HELPER.get_all_probes()
            
            self.la_cancel_edit()
        except Exception as ex:
            STATUS_MSG.append(f"Fehler beim Speichern: [{ex}]")
            self.feedback_message("error", STATUS_MSG)

    def la_edit_save(self, kennung: str, datum: str):
        """ Saves the edited Laborauswertung entry to the db

        Args:
            kennung (str): Kennung from the DB
                e.g.: 'ENE 1234'
            datum (str): Datum from the incom Probe
                e.g.: '12.12.2012
        """
        global STATUS_MSG, ALL_DATA_PROBE   

        try:
            DATABASE_HELPER.edit_laborauswertung(self.la_changed_item_lst, kennung, datum)
            STATUS_MSG = []
            self.feedback_message("success", ["Erfolgreich gespeichert"])
            # TODO: Aktualisiere die Tabelle, und ALL DATA PROBE
            ALL_DATA_PROBE = DATABASE_HELPER.get_all_probes()
            # TODO: Schließe die Edit Frame
            self.la_cancel_edit()
        except Exception as ex:
            STATUS_MSG.append(f"Fehler beim Speichern: [{ex}]")
            self.feedback_message("error", STATUS_MSG)

    def la_cancel_edit(self) -> None:
        """ Cancels the edit Process and hides the edit frame
        """
        global STATUS_MSG
        self.la_edit_frame_2.hide()
        STATUS_MSG = []



class Probe(QtWidgets.QMainWindow): 
    def __init__(self, parent=None):
        super(Probe, self).__init__(parent)
        uic.loadUi(r'./views/select_probe.ui', self)

        self.setWindowTitle(f"CapZa - Zasada - { __version__ } - Wähle Probe")

        self.load_probe_btn.clicked.connect(self.load_probe)
        self.cancel_btn.clicked.connect(self.close_window)
        self.init_shadow(self.load_probe_btn)
        self.init_shadow(self.cancel_btn)

        self.probe_filter_lineedit.textChanged.connect(self.filter_probe)

    def filter_probe(self, filter_text: str) -> None:
        """ Filter the Probe based on insert text
            TODO: Not working correctly

        Args:
            filter_text (str): User input text shall filter text
        """
        for i in range(self.tableWidget.rowCount()):
            item = self.tableWidget.item(i, 1)
            match = filter_text.lower() not in item.text().lower()
            self.tableWidget.setRowHidden(i, match)
            if not match:
                break

    def init_data(self, dataset: list[dict]) -> None:
        """ Inputs all the Probe data into the TableWidget

        Args:
            dataset (list[dict]): List of dictionaries with Probevalues
        """

        for row in dataset:
            rowPosition = self.tableWidget.rowCount()
            self.tableWidget.insertRow(rowPosition)
            self.tableWidget.setItem(rowPosition , 0, QTableWidgetItem(str(row["Datum"])))
            self.tableWidget.setItem(rowPosition , 1, QTableWidgetItem(str(row["Kennung"])))
            self.tableWidget.setItem(rowPosition , 2, QTableWidgetItem(str(row["Materialbezeichnung"])))

    def init_shadow(self, widget: QtWidgets) -> None:
        """ Adds shadow to the given widget

        Args:
            widget (QtWidgets): Widget to whom the shadow should be applied
                e.g.: QFrame, QPushButton, ...
        """
        effect = QGraphicsDropShadowEffect()
        effect.setOffset(0, 1)
        effect.setBlurRadius(8)
        widget.setGraphicsEffect(effect)  

    def load_probe(self) -> None:
        """ Gets the selected Probe and closes the Probe window
        """
        try:
            global SELECTED_PROBE
            row = self.tableWidget.currentRow()
            kennung = self.tableWidget.item(row,1).text()

            SELECTED_PROBE= DATABASE_HELPER.get_specific_probe(kennung)

            self.differentiate_probe(SELECTED_PROBE["Kennung"])
            self.parent().insert_values()
            self.close_window()
        except Exception as ex:
            STATUS_MSG.append(ex)
            self.parent().feedback_message("attention", STATUS_MSG)

    def differentiate_probe(self, wert: str) -> None:
        """ Decidees based on the wert where to look for information

        Args:
            wert (str): Number from the selected Probe
                e.g.: 'ENE123', '22-0000', ...

        Raises:
            Exception: Error when loading information to the probe
        """ 

        global ALL_DATA_NACHWEIS, STATUS_MSG, SELECTED_NACHWEIS
        try:
            if re.match("[0-9]+-[0-9]+", wert):
                ### es ist eine Projektnummer
                self.check_in_projekt_nummer(wert)
            elif re.match("[a-zA-Z]+\s[0-9]+", wert):
                ### es ist eine XXX 000 Kennung
                self.check_in_uebersicht_nachweis(self.get_full_project_ene_number(wert))
            elif re.match("[a-zA-Z]+\s['I']+", wert):
                SELECTED_NACHWEIS = 0
                raise Exception("DK Proben wurden nicht implementiert.")
            else:
                SELECTED_NACHWEIS = 0
                raise Exception("Andere Proben wurden nicht implementiert.")
        except Exception as ex:
            raise ex

    def get_full_project_ene_number(self, nummer: str) -> str:
        """ Gets the whole , VNE, ... Projectnr. from the shortform

        Args:
            nummer (str): Shortform of the ENE Nr.
                e.g.: ENE1234

        Returns:
            str: Entire ENE Nr.
                e.g.: ENE382981234 s
        """
        print("GF: ", nummer)

        letters, numbers = nummer.split()
        for index, nummer in ALL_DATA_NACHWEIS["Nachweisnr. Werk 1"].items():
            if isinstance(letters, str):
                if isinstance(numbers, str):
                    if isinstance(nummer, str):
                        if letters and numbers in nummer:
                            return letters, numbers, nummer
        else:
            return "/", "/"
            
    def check_in_uebersicht_nachweis(self, kennung_tpl: tuple) -> None:
        """ Loads the Nachweis data from Übersicht Nachweise

        Args:
            projektnummer (str): Nachweis Nr.
                e.g.: ('ENE', '2054', 'ENE5R3822054)
        """ 

        global SELECTED_NACHWEIS
        nachweis_data = ALL_DATA_NACHWEIS[ALL_DATA_NACHWEIS['Nachweisnr. Werk 1'] == kennung_tpl[2]]
        SELECTED_PROBE['Kennung_letters'] = kennung_tpl[0]
        SELECTED_PROBE['Kennung_nr'] = kennung_tpl[1]
        SELECTED_PROBE['Project_yr'] = "-"
        SELECTED_PROBE['Project_nr'] = "-"
        SELECTED_NACHWEIS = nachweis_data

    def check_in_projekt_nummer(self, projektnummer: str) -> None:
        """ Loads the Nachweis data from ProjektNr.

        Args:
            projektnummer (str): Project Nr.
                e.g.: "22-0000"
        """ 

        print("PN", projektnummer)

        global SELECTED_NACHWEIS
        projekt_data = ALL_DATA_PROJECT_NR[ALL_DATA_PROJECT_NR['Projekt-Nr.'] == str(projektnummer)]
        projekt_data["ORT"] = projekt_data["Ort"]
        projekt_data["PLZ"] = ""
        projekt_data["t"] = projekt_data["Menge [t/a]"]


        SELECTED_PROBE['Kennung_letters'] = "-"
        SELECTED_PROBE['Kennung_nr'] = "-"
        SELECTED_PROBE['Project_yr'] = projektnummer.split("-")[0]
        SELECTED_PROBE['Project_nr'] = projektnummer.split("-")[1]
        SELECTED_NACHWEIS = projekt_data

    def close_window(self) -> None:
        """ Closes the entire Window
        """
        self.close()

    def closeEvent(self, event):
        LOCKFILE_PROBE.unlock()
        

class Error(QtWidgets.QDialog): 
    def __init__(self, parent=None):
        super(Error, self).__init__(parent)
        uic.loadUi(r'./views/error.ui', self)
        global STATUS_MSG
        self.setWindowTitle(f"CapZa - Zasada - {__version__} - Fehlerbeschreibung")
        error_long_msg = "Es wurden mehrere Fehler gefunden:  "
        if len(STATUS_MSG) > 1:
            for error in STATUS_MSG:
                error_long_msg+=f"- {str(error)} "
        elif len(STATUS_MSG) == 1:
            error_long_msg = str(STATUS_MSG[0])
        else: 
            error_long_msg = "/"

        self.error_lbl.setText(error_long_msg)
        self.init_shadow(self.close_error_info_btn)
        self.init_shadow(self.error_msg_frame)
        self.close_error_info_btn.clicked.connect(self.delete_error)

    def init_shadow(self, widget: QtWidgets) -> None:
        """ Applies shadow to given Widget

        Args:
            widget (QtWidgets): Any QtWidget
                e.g.: QFrame, QPushButton,...
        """
        effect = QGraphicsDropShadowEffect()
        effect.setOffset(0, 1)
        effect.setBlurRadius(8)
        widget.setGraphicsEffect(effect)

    def close_window(self) -> None:
        """ Closes entire Window
        """

        self.close()
    
    def closeEvent(self, event):
        self.delete_error()
        LOCKFILE_ERROR.unlock()

    def delete_error(self) -> None:
        """ Deletes all errors
        """

        global STATUS_MSG
        STATUS_MSG = []
        self.error_lbl.setText("")
        self.close()
        self.parent()._check_for_errors()


if __name__ == "__main__":

    lockfile_main = QtCore.QLockFile(QtCore.QDir.tempPath() + 'capza.lock')

    if lockfile_main.tryLock(100):
        app = QtWidgets.QApplication(sys.argv)
        # Create and display the splash screen
        splash_pix = QPixmap("./assets/icon_logo.png")
        splash = QSplashScreen(splash_pix, Qt.WindowStaysOnTopHint)
        splash.setMask(splash_pix.mask())
        splash.show()
        

        d = {}
        try:
            d = CONFIG_HELPER.get_all_config()
        except Exception as ex:
            print(ex)
        if d:
            NW_PATH = d["nw_path"]
            PNR_PATH = d["project_nr_path"]
            STANDARD_SAVE_PATH =d["save_path"]
            DB_PATH = d["db_path"]
            
        try:
            ALL_DATA_NACHWEIS = pd.read_excel(NW_PATH)
        except Exception as ex:
            STATUS_MSG.append(f"Es wurde keine Nachweisliste gefunden. Bitte prüfe in den Referenzeinstellungen. [{str(ex)}]")

        try:
            ALL_DATA_PROBE = DATABASE_HELPER.get_all_probes()
        except Exception as ex:
            STATUS_MSG.append(f"Es konnten keine Proben geladen werden: [{ex}]")

        win = Ui()

        if STATUS_MSG != []:
            win.feedback_message("error", STATUS_MSG)
        
        win._check_for_errors()
        
        splash.finish(win)
        win.show()
        sys.exit(app.exec_())
    else: print("Running already")