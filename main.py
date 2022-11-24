
import datetime
from docx2pdf import convert
import pandas as pd
from PyQt5 import QtWidgets, uic, QtCore, QtGui
from PyQt5.QtGui import QIcon, QPixmap, QFont, QStandardItemModel, QStandardItem, QIntValidator
from PyQt5.QtCore import Qt, QDate, QTimer
from PyQt5.QtWidgets import QFileDialog, QTableWidgetItem, QGraphicsDropShadowEffect, QSplashScreen, QProgressDialog, QDateEdit, QHeaderView, QComboBox, QPushButton, QCommandLinkButton
import os, sys
import re

import assets.icons

from threading import Thread
import os
dirname = os.path.dirname(__file__)

# import helper modules
from modules.word_helper import Word_Helper
from modules.config_helper import ConfigHelper
from modules.db_helper.__init__ import DatabaseHelper

today_date_raw = datetime.datetime.now()
TODAY_FORMAT_STRING = today_date_raw.strftime(r"%d.%m.%Y")

LOCKFILE_PROBE = QtCore.QLockFile(QtCore.QDir.tempPath() + 'capza_probe.lock')
LOCKFILE_ERROR = QtCore.QLockFile(QtCore.QDir.tempPath() + 'capza_error.lock')


CONFIG_HELPER = ConfigHelper()
DATABASE_HELPER = 0

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
        self._ene = list()
        self._pnr = list()
        self.showMaximized()
        self.init_main()

    def init_main(self) -> None:
        # STANDARDEINSTELLUNGEN:
        self.word_helper = Word_Helper()
        self.setWindowTitle(f"CapZa - Zasada - { __version__ } ")
        self.setWindowIcon(QIcon(r'./assets/icon_logo.png'))
        self.stackedWidget.setCurrentIndex(0)

        self.logo_right_lbl.setPixmap(QPixmap("./assets/l_logo.png"))
        self.second_info_lbl.hide()
        self.main_version_lbl.setText(__version__)
        self.error_info_btn.clicked.connect(self.showError)
        self.status_msg_btm.hide()
        self.hide_admin_msg_btn.clicked.connect(self.hide_admin_msg)
        
        # NAVIGATION:
        self.nav_data_btn.clicked.connect(lambda : self.display(0))
        self.nav_analysis_btn.clicked.connect(lambda : self.display(1))
        self.nav_einstufung_btn.clicked.connect(lambda: self.display(2))
        self.nav_pnp_entry_btn.clicked.connect(lambda : self.display(3))
        self.nav_pnp_output_btn.clicked.connect(lambda : self.display(4))
        self.nav_order_form_btn.clicked.connect(lambda : self.display(5))
        self.nav_settings_btn.clicked.connect(lambda : self.display(6))
        self.nav_laborauswertung_btn.clicked.connect(lambda : self.display(7))

        init_shadow(self.data_1)
        init_shadow(self.data_1_2)  
        init_shadow(self.data_2)
        init_shadow(self.data_3)
        init_shadow(self.select_probe_btn)
        init_shadow(self.migrate_btn)
        init_shadow(self.aqs_btn)
        init_shadow(self.pnp_out_empty)
        init_shadow(self.pnp_out_protokoll_btn)
        init_shadow(self.auftrag_empty)
        init_shadow(self.auftrag_letsgo)
        init_shadow(self.analysis_f1)
        init_shadow(self.analysis_f2)
        init_shadow(self.pnp_o_frame)
        init_shadow(self.order_frame)
        init_shadow(self.pnp_in_allg_frame)
        init_shadow(self.clear_cache_btn)   
        init_shadow(self.vor_ort_frame)     

        ### DATENEINGABE:
        self.end_dateedit.setDate(self._get_todays_date_qdate())
        self.load_project_nr()
        self.select_probe_btn.clicked.connect(self.read_all_probes)
        self.brandtest_combo.currentTextChanged.connect(lambda: self.analysis_brandtest_lineedit.setText(self.brandtest_combo.currentText()))
        self.nh3_lineedit_2.textChanged.connect(self.check_nh3_value)
        self.h2_lineedit_2.textChanged.connect(lambda: self.h2_lineedit.setText(self.h2_lineedit_2.text()))
        self.kennung_rb.clicked.connect(lambda: self.empty_manual_search(self.pnr_combo, self.project_nr_lineedit))
        self.pnr_rb.clicked.connect(lambda: self.empty_manual_search(self.kennung_combo, self.kennung_lineedit))

        rx = QtCore.QRegExp("\d+")
        self.kennung_lineedit.setMaxLength(4)
        self.kennung_lineedit.setValidator(QtGui.QRegExpValidator(rx))
        self.project_nr_lineedit.setMaxLength(4)
        self.project_nr_lineedit.setValidator(QtGui.QRegExpValidator(rx))
        self.search_manually_btn.clicked.connect(self.search_manual)

        try:
            self.kennung_combo.addItems(self.extract_all_ene_values())
            self.pnr_combo.addItems(self.extract_all_pnr_years())
        except Exception as ex:
            print(ex)


        ### ANALYSEWERTE:
        self.migrate_btn.clicked.connect(self.create_bericht_document)
        self.aqs_btn.clicked.connect(self._no_function)

        ### PNP Input:
        self.pnp_in_create_protokoll.clicked.connect(self.trigger_pnp_in)

        ### PNP Output:
        int_validator = QIntValidator(0, 999999999, self)
        self.output_nr_lineedit.setValidator(int_validator)
        self.pnp_out_protokoll_btn.clicked.connect(self.create_pnp_out_protokoll)

        ### AUFTRAGSFORMULAR:
        self.autrag_load_column_view()
        self.auftrag_add_auftrag_btn.clicked.connect(self.auftrag_add_auftrag)

        ### LABORAUSWERTUNG:
        self.la_table_view.doubleClicked.connect(self.edit_laborauswertung)
        self.add_laborauswertung_btn.clicked.connect(self.add_laborauswertung)
        self.la_table_view.horizontalHeader().setStyleSheet("QHeaderView {font-size: 12pt; font-weight:bold} QTableView {border: 1px solid black;}")
        self.la_table_view.resizeColumnsToContents()
        init_shadow(self.la_table_view)

        ### GRENZWERTEINSTELLUNGEN:
        self.grenz_help_frame.hide()
        self.grenz_help_frame_shown = False
        self.grenz_toggle_help_btn.clicked.connect(self.grenz_toggle_help_frame)

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

        self.clear_cache_btn.clicked.connect(self._clear_cache)
        self.save_references_btn.clicked.connect(self.save_references)
        
        self.check_la_enable()
        self.check_la_db_path.toggled.connect(self.check_la_enable)
        
        if self.nw_overview_path.text() == "" or self.project_nr_path.text()=="":
            STATUS_MSG.append("Es sind keine Nachweise hinterlegt. Prüfe in den Referenzeinstellungen.")
            self.feedback_message("error", STATUS_MSG)

    def check_nh3_value(self) -> None:
        self.nh3_lineedit.setText(self.nh3_lineedit_2.text())
        if float(self.nh3_lineedit.text()) <= 20:
            self.set_ampel_color(self.nh3_ampel_lbl, "green")
        elif float(self.nh3_lineedit.text()) > 20:
            self.set_ampel_color(self.nh3_ampel_lbl, "red")

    def hide_admin_msg(self) -> None:
        self.admin_msg_frame.hide()

    def _clear_cache(self) -> None:
        """ Resets all standard variables to its default
        """
        global SELECTED_PROBE,SELECTED_NACHWEIS,ALL_DATA_PROBE,ALL_DATA_NACHWEIS,ALL_DATA_PROJECT_NR,NW_PATH,PNR_PATH,STATUS_MSG,BERICHT_FILE,ALIVE, PROGRESS
        SELECTED_PROBE = 0
        SELECTED_NACHWEIS = 0
        NW_PATH = ""
        PNR_PATH = ""
        STATUS_MSG = []
        BERICHT_FILE = ""

    def _get_todays_date_qdate(self) -> QDate:
        """ Get QDate object from current date string

        Returns:
            QDate: Contains the current Date
        """

        d,m,y = TODAY_FORMAT_STRING.split(".")
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

    def trigger_pnp_in(self) -> bool:
        """ Opens a dialog after setting the pnp in-
        """
        dlg = QtWidgets.QMessageBox(self)
        dlg.setWindowTitle("PNP Input mit oder ohne Weiterberechnungsformular")
        if self.weiterberechnung_checkBox.checkState() == 2:
            dlg.setText("Soll das Weiterberechnungsformular wirklich dazu erstellt werden?")
        else:
            dlg.setText("Soll wirklich kein Weiterberechnungsformular dazu erstellt werden? [Wenn nicht, dann wähle 'No']s")
        dlg.setStandardButtons(QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
        dlg.setIcon(QtWidgets.QMessageBox.Question)
        button = dlg.exec()

        if button == QtWidgets.QMessageBox.Yes:
            print("Yes!")
            return True
        else:
            self.create_weiterberechnung_form()
            print("No!")
            return False    

    def create_pnp_in(self):
        if self.trigger_pnp_in():
            # Create file
            print("Crreate file")
        pass

    def create_weiterberechnung_form(self):
        pass
        
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
                    SELECTED_PROBE = DATABASE_HELPER.get_specific_probe(projectnr)
                    for sheet in ALL_DATA_PROJECT_NR:
                        try:
                            if str(projectnr) in ALL_DATA_PROJECT_NR[sheet]["Projekt-Nr."].values:
                                nachweis_data = ALL_DATA_PROJECT_NR[sheet][ALL_DATA_PROJECT_NR[sheet]['Projekt-Nr.'] == str(projectnr)]
                        except Exception as ex:
                            STATUS_MSG = [f"Probe mit Sheet {sheet} konnte nicht geladen werden: [{ex}]"]
                            self.feedback_message("error", STATUS_MSG)
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
            STATUS_MSG = [f"Fehler: {str(ex)}"]
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

    ### Grenzwerteinstellungen
    def grenz_toggle_help_frame(self):
        if self.grenz_help_frame_shown == True:
            self.grenz_help_frame.hide()
            self.grenz_help_frame_shown = False
        else:
            self.grenz_help_frame.show()
            self.grenz_help_frame_shown = True

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
            ALL_DATA_PROJECT_NR = pd.read_excel(PNR_PATH, sheet_name=None)
            STATUS_MSG = []
        except Exception as ex:
            STATUS_MSG.append(f"Projektnummern konnten nicht geladen werden: [{ex}]")
            self.feedback_message("error", STATUS_MSG)

    def showError(self) -> None:
        """ Shows the error frame
        """
        if LOCKFILE_ERROR.tryLock(100):
            error = Error(self)
            error.show()
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
        global DATABASE_HELPER
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
                DATABASE_HELPER = DatabaseHelper(DB_PATH)
                ALL_DATA_PROBE = DATABASE_HELPER.get_all_probes()

            references = {
                "nw_path": nw_path,
                "project_nr_path": project_nr_path,
                "save_path": save_path,
                "la_path": la_path,
                "db_path": db_path
            }

            for key, value in references.items():
                CONFIG_HELPER.update_specific_value(key, value)

            self.laborauswertung_path.setText("")

            self.feedback_message("success", ["Neue Referenzen erfolgreich gespeichert."])
            self.init_main()
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
    
        if i == 6:
            self.show_second_info("Der Pfad zur 'Nachweis Übersicht' Excel ist nur temporär und wird in Zukunft durch Echtdaten aus RAMSES ersetzt.")

        if i == 7:
            t1 = Thread(target=self.load_laborauswertung)
            t2 = Thread(target=self.feedback_message, args=("info", ["Laborauswertung wird geladen..."],))
            t1.start()
            t2.start()

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
            toc = round_if_psbl(float(self.toc_lineedit.text()))
        else:
            toc = ""

        if not SELECTED_PROBE["EC %"] == None:
            ec = round_if_psbl(float(self.ec_lineedit.text()))
        else:
            ec = ""
        
        if toc != "" and ec != "":
            toc_float = float(toc) 
            ec_float = float(ec)
            aoc = round_if_psbl(toc_float-ec_float)
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
            "heute": str(TODAY_FORMAT_STRING),
            "datum": date,
            #
            "wert":  round_if_psbl(float(self.ph_lineedit.text())) if self.ph_lineedit.text() != "-" else "",
            "leitfaehigkeit":  round_if_psbl(float(self.leitfaehigkeit_lineedit.text())) if self.leitfaehigkeit_lineedit.text() != "-" else "",
            "doc": round_if_psbl(float(self.doc_lineedit.text())) if self.doc_lineedit.text() != "-" else "",
            "molybdaen": round_if_psbl(float(self.mo_lineedit.text())) if self.mo_lineedit.text() != "-" else "",

            "selen": round_if_psbl(float(self.se_lineedit.text())) if self.se_lineedit.text() != "-" else "",
            "antimon": round_if_psbl(float(self.sb_lineedit.text())) if self.sb_lineedit.text() != "-" else "",
            "chrom": round_if_psbl(float(self.chrome_vi_lineedit.text())) if self.chrome_vi_lineedit.text() != "-" else "",
            "tds": round_if_psbl(float(self.tds_lineedit.text())) if self.tds_lineedit.text() != "-" else "",
            "chlorid": self.chlorid_lineedit.text() if self.chlorid_lineedit.text() != "-" else "",
            "fluorid": self.fluorid_lineedit.text() if self.fluorid_lineedit.text() != "-" else "",
            "feuchte": self.feuchte_lineedit.text() if self.feuchte_lineedit.text() != "-" else "",
            "lipos_ts": round_if_psbl(float(self.lipos_ts_lineedit.text())) if self.lipos_ts_lineedit.text() != "-" else "",
            "lipos_os": round_if_psbl(float(self.lipos_os_lineedit.text())) if self.lipos_os_lineedit.text() != "-" else "",
            "gluehverlust": round_if_psbl(float(self.gluehverlust_lineedit.text())) if self.gluehverlust_lineedit.text() != "-" else "",
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

        print(data["tds"])

        word_file = self.create_word(CONFIG_HELPER.get_specific_config_value("bericht_vorlage"), data, "Bericht")    
        if word_file != "/":    
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
        _date_edit = QDateEdit(self.auftrag_table_view, calendarPopup=True)
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

        probe_date = self.pnp_output_probenahmedatum.date().toString("dd.MM.yyyy")

        data = {
            "datum": probe_date,
            "probenehmer": probenehmer,
            "anwesende_personen": anwesende_personen,
            "output_nr": self.output_nr_lineedit.text(),
            "output_nr_1": str(int(self.output_nr_lineedit.text())+1),
            "output_nr_2": str(int(self.output_nr_lineedit.text())+2),
            "output_nr_3": str(int(self.output_nr_lineedit.text())+3),
            "output_nr_4": str(int(self.output_nr_lineedit.text())+4)
        }
        word_file= self.create_word(vorlage_document, data, "PNP Output Protokoll")
        try:
            thread1 = Thread(target=self.word_helper.open_word, args=(word_file,))
            thread2 = Thread(target=self.feedback_message, args=("info", ["Word wird geöffnet..."]))
            thread1.start() 
            thread2.start()
        except Exception as ex:
            self.feedback_message("attention", [f"Fehler beim Erstellen der Word Datei [{ex}]"])
            STATUS_MSG.append(str(ex))

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
                self.feedback_message("success", ["Das Protokoll wurde erfolgreich erstellt."])
                return file[0]
            else: 
                return "/"
        except Exception as ex:
            STATUS_MSG= [f"{dialog_file} konnte nicht erstellt werden: " + str(ex)]
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
            toc = round_if_psbl(float(SELECTED_PROBE["TOC %"]))
        else:
            toc = ""

        if not SELECTED_PROBE["EC %"] == "":
            ec = round_if_psbl(float(SELECTED_PROBE["EC %"]))
        else:
            ec = ""
        
        if isinstance(toc, float) and isinstance(ec, float):
            aoc = round_if_psbl(toc-ec)
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
                "leitfaehigkeit": str(SELECTED_PROBE["Leitfähigkeit mS/cm"]),
                "doc": round_if_psbl(SELECTED_PROBE["Bezogen auf das eingewogene Material DOC mg/L"]),
                "molybdaen": round_if_psbl(SELECTED_PROBE["Bezogen auf das eingewogene Material DOC mg/L"]),
                "selen": round_if_psbl(SELECTED_PROBE["Se 196.090 (Aqueous-Axial-iFR)"]),
                "antimon": round_if_psbl(SELECTED_PROBE["Sb 206.833 (Aqueous-Axial-iFR)"]),
                "chrom": round_if_psbl(SELECTED_PROBE["Cr 205.560 (Aqueous-Axial-iFR)"]),
                "tds": round_if_psbl(SELECTED_PROBE["TDS Gesamt gelöste Stoffe mg/L"]),
                "chlorid": str(SELECTED_PROBE["Chlorid mg/L"]),
                "fluorid": str(SELECTED_PROBE["Fluorid mg/L"]),
                "feuchte": str(SELECTED_PROBE["Wassergehalt %"]),
                "lipos_ts": round_if_psbl(SELECTED_PROBE["Lipos TS %"]),
                "lipos_os": round_if_psbl(SELECTED_PROBE["Lipos FS %"]),
                "gluehverlust": round_if_psbl(SELECTED_PROBE["GV %"]),
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
        self.gluehverlust_lineedit.setText("")
        self.doc_lineedit.setText("")
        self.tds_lineedit.setText("")
        self.mo_lineedit.setText("")
        self.se_lineedit.setText("")
        self.sb_lineedit.setText("")
        self.fluorid_lineedit.setText("")
        self.chlorid_lineedit.setText("")
        self.toc_lineedit.setText("")
        self.ec_lineedit.setText("")
        self.pb_lineedit.setText("")

        self.set_ampel_color(self.ph_ampel_lbl, "default")
        self.set_ampel_color(self.lf_ampel_lbl, "default")
        self.set_ampel_color(self.feuchte_ampel_lbl, "default")
        self.set_ampel_color(self.chrom_ampel_lbl, "default")
        self.set_ampel_color(self.ts_ampel_lbl, "default")
        self.set_ampel_color(self.os_ampel_lbl, "default")
        self.set_ampel_color(self.GV_ampel_lbl, "default")
        self.set_ampel_color(self.doc_ampel_lbl, "default")
        self.set_ampel_color(self.tds_ampel_lbl, "default")
        self.set_ampel_color(self.mo_ampel_lbl, "default")
        self.set_ampel_color(self.se_ampel_lbl, "default")
        self.set_ampel_color(self.sb_ampel_lbl, "default")
        self.set_ampel_color(self.fluorid_ampel_lbl, "default")
        self.set_ampel_color(self.chlorid_ampel_lbl, "default")
        self.set_ampel_color(self.toc_ampel_lbl, "default")
        self.set_ampel_color(self.ec_ampel_lbl, "default")
        self.set_ampel_color(self.nh3_ampel_lbl, "default")
        self.set_ampel_color(self.h2_ampel_lbl, "default")
        self.set_ampel_color(self.brandtest_ampel_lbl, "default")
        self.set_ampel_color(self.pb_ampel_lbl, "default")

    def insert_values(self) -> None:
        """ Inserts all value into CapZa FE based on selected Pobe
        """

        global SELECTED_PROBE
        global SELECTED_NACHWEIS
        global STATUS_MSG

        self.empty_values()
        STATUS_MSG = []
        ### in Dateneingabe
        try:
            self.kennung_combo.setCurrentText(format_specific_insert_value("Kennung_letters", SELECTED_PROBE))
            self.pnr_combo.setCurrentText(format_specific_insert_value("Project_yr", SELECTED_PROBE))
            self.kennung_lineedit.setText(format_specific_insert_value("Kennung_nr", SELECTED_PROBE))
            self.project_nr_lineedit.setText(format_specific_insert_value("Project_nr", SELECTED_PROBE))
            self.name_lineedit.setText(format_specific_insert_value("Material", SELECTED_NACHWEIS))
            self.person_lineedit.setText(format_specific_insert_value("Erzeuger", SELECTED_NACHWEIS))
            self.location_lineedit.setText(str(list(SELECTED_NACHWEIS["PLZ"])[0]) + " " + str(list(SELECTED_NACHWEIS["ORT"])[0]))
            self.avv_lineedit.setText(self.format_avv_space_after_every_second(str(list(SELECTED_NACHWEIS["AVV"])[0])))
            self.amount_lineedit.setText("{:,}".format(list(SELECTED_NACHWEIS["t"])[0]).replace(",", "."))
        except Exception as ex:
            print("FEHLER = " + str(ex))
        ### in Analysewerte
        try:
            self.ph_lineedit.setText(str(round_if_psbl(format_specific_insert_value("pH-Wert", SELECTED_PROBE))))
            try:
                if float(SELECTED_PROBE["pH-Wert"]) <= 8:
                    self.set_ampel_color(self.ph_ampel_lbl, "yellow")
            except: pass
        except Exception as ex:
            STATUS_MSG.append(f"Der pH-Wert kann nicht interpretiert werden: [{ex}]")
        try:
            self.leitfaehigkeit_lineedit.setText(str(round_if_psbl(format_specific_insert_value("Leitfähigkeit mS/cm", SELECTED_PROBE))))
            try:
                if float(SELECTED_PROBE["Leitfähigkeit mS/cm"]) >= 12:
                    self.set_ampel_color(self.lf_ampel_lbl, "yellow")
            except:
                pass
        except Exception as ex:
            STATUS_MSG.append(f"Der Leitfähigkeitswert [mS/cm] kann nicht interpretiert werden: [{ex}]")
        try:
            self.feuchte_lineedit.setText(str(round_if_psbl(format_specific_insert_value("Wassergehalt %", SELECTED_PROBE))))
        except Exception as ex:
            STATUS_MSG.append(f"Der Wassergehalt [%] kann nicht interpretiert werden: [{ex}]")
        
        try:
            self.chrome_vi_lineedit.setText(str(round_if_psbl(format_specific_insert_value("Cr 205.560 (Aqueous-Axial-iFR)", SELECTED_PROBE))))
            try:
                if float(SELECTED_PROBE["Cr 205.560 (Aqueous-Axial-iFR)"]) <= 7:
                    self.set_ampel_color(self.chrom_ampel_lbl, "green")
                elif float(SELECTED_PROBE["Cr 205.560 (Aqueous-Axial-iFR)"]) > 7:
                    self.set_ampel_color(self.chrom_ampel_lbl, "red")
            except: pass
        except Exception as ex:
            STATUS_MSG.append(f"Der Chromwert kann nicht interpretiert werden: [{ex}]")
        try:
            self.lipos_ts_lineedit.setText(str(round_if_psbl(format_specific_insert_value("Lipos TS %", SELECTED_PROBE))))
            try:
                if float(SELECTED_PROBE["Lipos TS %"]) <= 4:
                    self.set_ampel_color(self.ts_ampel_lbl, "green")
                elif float(SELECTED_PROBE["Lipos TS %"]) > 4:
                    self.set_ampel_color(self.ts_ampel_lbl, "red")
            except:pass
        except Exception as ex:
            STATUS_MSG.append(f"Der Lipos TS [%] kann nicht interpretiert werden: [{ex}]")
        try:
            self.lipos_os_lineedit.setText(format_specific_insert_value("Lipos FS %", SELECTED_PROBE))
        except Exception as ex:
            STATUS_MSG.append(f"Der Lipos OS [%] kann nicht interpretiert werden: [{ex}]")
        try:
            self.gluehverlust_lineedit.setText(str(round_if_psbl(format_specific_insert_value("GV %", SELECTED_PROBE))))
            try:
                if float(SELECTED_PROBE["GV %"]) <= 10:
                    self.set_ampel_color(self.GV_ampel_lbl, "green")
                elif float(SELECTED_PROBE["GV %"]) > 10:
                    self.set_ampel_color(self.GV_ampel_lbl, "red")
            except: pass
        except Exception as ex:
            STATUS_MSG.append(f"Der Glühverlust-Wert [%] kann nicht interpretiert werden: [{ex}]")
        try:
            self.doc_lineedit.setText(str(round_if_psbl(format_specific_insert_value("Bezogen auf das eingewogene Material DOC mg/L", SELECTED_PROBE))))
            try:
                if float(SELECTED_PROBE["Bezogen auf das eingewogene Material DOC mg/L"]) <= 100:
                    self.set_ampel_color(self.doc_ampel_lbl, "green")
                elif 100 < float(SELECTED_PROBE["Bezogen auf das eingewogene Material DOC mg/L"]) <= 199:
                    self.set_ampel_color(self.doc_ampel_lbl, "purple")
                elif float(SELECTED_PROBE["Bezogen auf das eingewogene Material DOC mg/L"]) > 199:
                    self.set_ampel_color(self.doc_ampel_lbl, "red")
            except: pass
        except Exception as ex:
            STATUS_MSG.append(f"Der DOC-Wert [mg/L] kann nicht interpretiert werden: [{ex}]")
        try:
            self.tds_lineedit.setText(str(round_if_psbl(format_specific_insert_value("TDS Gesamt gelöste Stoffe mg/L", SELECTED_PROBE))))
            try:
                if float(SELECTED_PROBE["TDS Gesamt gelöste Stoffe mg/L"]) <= 10000:
                    self.set_ampel_color(self.tds_ampel_lbl, "green")
                elif 10000 < float(SELECTED_PROBE["TDS Gesamt gelöste Stoffe mg/L"]) < 20000:
                    self.set_ampel_color(self.tds_ampel_lbl, "purple")
                elif float(SELECTED_PROBE["TDS Gesamt gelöste Stoffe mg/L"]) > 20000:
                    self.set_ampel_color(self.tds_ampel_lbl, "red")
            except: pass
        except Exception as ex:
            STATUS_MSG.append(f"Der Wert 'TDS Gesamt gelöste Stoffe (mg/L)' kann nicht interpretiert werden: [{ex}]")
        try:            
            self.mo_lineedit.setText(str(round_if_psbl(format_specific_insert_value("Bezogen auf das eingewogene Material Molybdän mg/L", SELECTED_PROBE))))
            try:
                if float(SELECTED_PROBE["Bezogen auf das eingewogene Material Molybdän mg/L"]) <= 3:
                    self.set_ampel_color(self.mo_ampel_lbl, "green")
                elif float(SELECTED_PROBE["Bezogen auf das eingewogene Material Molybdän mg/L"]) > 3:
                    self.set_ampel_color(self.mo_ampel_lbl, "red")
            except: pass
        except Exception as ex:
            STATUS_MSG.append(f"Der Molybdän-Wert kann nicht interpretiert werden: [{ex}]")
        try:
            self.se_lineedit.setText(str(round_if_psbl(format_specific_insert_value("Se 196.090 (Aqueous-Axial-iFR)", SELECTED_PROBE))))
            try:
                if float(SELECTED_PROBE["Se 196.090 (Aqueous-Axial-iFR)"]) <= 0.7:
                    self.set_ampel_color(self.se_ampel_lbl, "green")
                elif float(SELECTED_PROBE["Se 196.090 (Aqueous-Axial-iFR)"]) > 0.7:
                    self.set_ampel_color(self.se_ampel_lbl, "red")
            except: pass
        except Exception as ex:
            STATUS_MSG.append(f"Der Se-Wert kann nicht interpretiert werden: [{ex}]")
        try:
            self.sb_lineedit.setText(str(round_if_psbl(format_specific_insert_value("Sb 206.833 (Aqueous-Axial-iFR)", SELECTED_PROBE))))
            try:
                if float(SELECTED_PROBE["Sb 206.833 (Aqueous-Axial-iFR)"]) <= 0.5:
                    self.set_ampel_color(self.sb_ampel_lbl, "green")
                elif float(SELECTED_PROBE["Sb 206.833 (Aqueous-Axial-iFR)"]) > 0.7:
                    self.set_ampel_color(self.sb_ampel_lbl, "red")
            except: pass
        except Exception as ex:
            STATUS_MSG.append(f"Der Sb-Wert kann nicht interpretiert werden: [{ex}]")
        try:
            self.fluorid_lineedit.setText(str(round_if_psbl(format_specific_insert_value("Fluorid mg/L", SELECTED_PROBE))))
            try:
                if float(SELECTED_PROBE["Fluorid mg/L"]) <= 50:
                    self.set_ampel_color(self.fluorid_ampel_lbl, "green")
                elif float(SELECTED_PROBE["Fluorid mg/L"]) > 50:
                    self.set_ampel_color(self.fluorid_ampel_lbl, "red")
            except: pass
        except Exception as ex:
            STATUS_MSG.append(f"Der Fluorid-Wert [mg/L] kann nicht interpretiert werden: [{ex}]")
        try:
            self.chlorid_lineedit.setText(str(round_if_psbl(format_specific_insert_value("Chlorid mg/L", SELECTED_PROBE))))
            try:
                if float(SELECTED_PROBE["Chlorid mg/L"]) <= 2500:
                    self.set_ampel_color(self.chlorid_ampel_lbl, "green")
                elif 2500 < float(SELECTED_PROBE["Chlorid mg/L"]) <= 4000:
                    self.set_ampel_color(self.chlorid_ampel_lbl, "purple")
                elif float(SELECTED_PROBE["Chlorid mg/L"]) > 4000:
                    self.set_ampel_color(self.chlorid_ampel_lbl, "red")
            except: pass
        except Exception as ex:
            STATUS_MSG.append(f"Der Chlorid-Wert [mg/L] kann nicht interpretiert werden: [{ex}]")
        try:
            self.toc_lineedit.setText(str(round_if_psbl(format_specific_insert_value("TOC %", SELECTED_PROBE))))
            try:
                if float(SELECTED_PROBE["TOC %"]) <= 6:
                    self.set_ampel_color(self.toc_ampel_lbl, "green")
                elif float(SELECTED_PROBE["TOC %"]) > 6:
                    self.set_ampel_color(self.toc_ampel_lbl, "red")
            except: pass
        except Exception as ex:
            STATUS_MSG.append(f"Der TOC-Wert [%] kann nicht interpretiert werden: [{ex}]")
        try:
            self.ec_lineedit.setText(str(round_if_psbl(format_specific_insert_value("EC %", SELECTED_PROBE))))
        except Exception as ex:
            STATUS_MSG.append(f"Der EC-Wert [%] kann nicht interpretiert werden: [{ex}]")
        try:
            if SELECTED_PROBE["Pb"] != "<LOD":
                self.pb_lineedit.setText(str(float(SELECTED_PROBE["Pb"]) * 10000)) if SELECTED_PROBE["Pb"] != None else "-"
            else:
                self.pb_lineedit.setText("-")
        except Exception as ex:
            STATUS_MSG.append(f"Der Pb kann nicht interpretiert werden: [{ex}]")
        
        date = str(SELECTED_PROBE["Datum"]).split()[0]
        date = date.split("-")
        y = date[0]
        m = date[1]
        d = date[2]
        self.probe_date.setDate(QDate(int(y), int(m), int(d)))
        self.check_start_date.setDate(QDate(int(y),int(m),int(d)))
        self.pnp_output_probenahmedatum.setDate(QDate(int(y), int(m), int(d)))
        self.pnp_input_date_edit.setDate(QDate(int(y), int(m), int(d)))


        self.nachweisnr_lineedit.setText(str(SELECTED_PROBE["Kennung"]))

        ### in PNP Input
        self.pnp_in_erzeuger_lineedit.setText(str(list(SELECTED_NACHWEIS["Erzeuger"])[0]))
        self.pnp_in_abfallart_textedit.setPlainText(self.format_avv_space_after_every_second(str(list(SELECTED_NACHWEIS["AVV"])[0])) + ", " + str(list(SELECTED_NACHWEIS["Material"])[0]) )


        if len(STATUS_MSG) > 0:
            self.feedback_message("attention", ["Ein oder mehr Werte konnten nicht interpretiert werden."])
        else:
            self.feedback_message("success", ["Probe erfolgreich geladen."])
        self.show_second_info("Gehe zu 'Analysewerte', um die Dokumente zu erstellen. >")         

    def set_ampel_color(self, ampel_lbl: QtWidgets.QLabel, color: str) -> None:
        """ Sets the color of a label (Ampel)

        Args:
            ampel_lbl (QtWidgets.QLabel): Label whos color will get changeds
            color (str): Color that will be applied to the Label
        """

        if color == "green":
            color = "#16de29"
        elif color == "red":
            color = "#fa1b1b"
        elif color == "yellow":
            color = "#faec1d"
        elif color == "purple":
            color = "#b700d3"
        else: color = "#ffffff"

        ampel_lbl.setStyleSheet(
        "QLabel { "
            f"background: {color}"
        "}"
        )

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

    def insert_values_in_la_table(self) -> None:
        global ALL_DATA_PROBE

        self.model = QtGui.QStandardItemModel(len(ALL_DATA_PROBE), 13)
        self.model.setHorizontalHeaderLabels(["Datum","Kennung", "Materialbezeichnung", "TS [%]", "Wasserfaktor [%]", "Wasserfaktor (getrocknet) [%]", "Lipos TS [%]", "Lipos FS [%]","Lipos aus Frischsubstanz [%]","GV [%]","TDS Gesamt gelöste Stoffe [mg/L]","Einwaage TS","Faktor"])
        for row, text in enumerate(ALL_DATA_PROBE):
            date_item = QtGui.QStandardItem(str(text["Datum"]) if text["Datum"] != None else "")
            kennung_item = QtGui.QStandardItem(str(text["Kennung"]) if text["Kennung"] != None else "")
            material_item = QtGui.QStandardItem(str(text["Materialbezeichnung"]) if text["Materialbezeichnung"] != None else "")
            ts_item = QtGui.QStandardItem(str(round_if_psbl(text["% TS"])) if text["% TS"] != None else "")
            wasser_item = QtGui.QStandardItem(str(round_if_psbl(text["Wasserfaktor"])) if text["Wasserfaktor"] != None else "")
            dry_wasser_item = QtGui.QStandardItem(str(round_if_psbl(text["Wasserfaktor getrocknetes Material"])) if text["Wasserfaktor getrocknetes Material"] != None else "")
            liposts_item = QtGui.QStandardItem(str(round_if_psbl(text["Lipos TS %"])) if text["Lipos TS %"] != None else "")
            liposfs_item = QtGui.QStandardItem(str(round_if_psbl(text["Lipos FS %"])) if text["Lipos FS %"] != None else "")
            liposfrisch_item = QtGui.QStandardItem(str(round_if_psbl(text[r"% Lipos ermittelts aus Frischsubstanz"])) if text["% Lipos ermittelts aus Frischsubstanz"] != None else "")
            gv_item = QtGui.QStandardItem(str(round_if_psbl(text["GV %"])) if text["GV %"] != None else "")
            tds_gesamt_item = QtGui.QStandardItem(str(round_if_psbl(text["TDS Gesamt gelöste Stoffe mg/L"])) if text["TDS Gesamt gelöste Stoffe mg/L"] != None else "")
            einwaage_ts_item = QtGui.QStandardItem(str(round_if_psbl(text["Einwaage TS"])) if text["Einwaage TS"] != None else "")
            faktor_item = QtGui.QStandardItem(str(round_if_psbl(text["Faktor"])) if text["Faktor"] != None else "")
            
            self.model.setItem(row, 0, date_item)
            self.model.setItem(row, 1, kennung_item)
            self.model.setItem(row, 2, material_item)
            self.model.setItem(row, 3, ts_item)
            self.model.setItem(row, 4, wasser_item)
            self.model.setItem(row, 5, dry_wasser_item)
            self.model.setItem(row, 6, liposts_item)
            self.model.setItem(row, 7, liposfs_item)
            self.model.setItem(row, 8, liposfrisch_item)
            self.model.setItem(row, 9, gv_item)
            self.model.setItem(row, 10, tds_gesamt_item)
            self.model.setItem(row, 11, einwaage_ts_item)
            self.model.setItem(row, 12, faktor_item)
        
        #filter proxy model
        self.filter_proxy_model = QtCore.QSortFilterProxyModel()
        self.filter_proxy_model.setSourceModel(self.model)
        self.filter_proxy_model.setFilterKeyColumn(1) # second column
        self.la_table_view.setModel(self.filter_proxy_model)
        self.la_search_found_lbl.setText(f"{self.filter_proxy_model.rowCount()} Ergebnisse gefunden")

        self.laborauswertung_lineedit.textChanged.connect(self.apply_filter)

    def apply_filter(self, text):
        self.filter_proxy_model.setFilterRegExp(text)
        self.la_search_found_lbl.setText(f"{self.filter_proxy_model.rowCount()} Ergebnisse gefunden")

    def load_laborauswertung(self) -> bool:
        """ Loads the whole Laborauswertung (from db) into the FE

        Raises:
            ex: When the Probes cannot be loaded

        Returns:
            bool: True when data is being loaded; False when there is an error
        """
        global ALL_DATA_PROBE
        self._ene = []
        self._pnr = []
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
            self.insert_values_in_la_table()
        except Exception as ex:
            STATUS_MSG.append(f"Fehler beim Laden der Probedaten: [{str(ex)}]")
            self.feedback_message("error", STATUS_MSG)
            return False

    # def insert_new_in_laborauswertung(self, row, col, value) -> None:
    #     self.model.setData(self.model.index(row, col, value))

    def edit_laborauswertung(self) -> None:
        """ Loads the selected Laborauswertung row and open the frame to edit
        """
        row=(self.la_table_view.selectionModel().currentIndex())
        # kennung=index.sibling(index.row(),1).data()
        try:
            selected_datum = row.sibling(row.row(),0).data()
            selected_kennung = row.sibling(row.row(),1).data()
            selected_material = row.sibling(row.row(),2).data()

            data = {
                "datum": selected_datum,
                "kennung": selected_kennung,
                "material": selected_material
            }

            la = Laborauswertung(self, "edit")
            la.insert_values(data)
            la.show()
        except Exception as ex:
            STATUS_MSG = [f"Das Datum, die Kennung und das Material müssen zwingend angegeben sein: [{ex}]"]
            self.feedback_message("error", STATUS_MSG)
        
    def add_laborauswertung(self) -> None:
        """ Open the new Windows
        """
        la = Laborauswertung(self, "add")
        la.show()

    def show_element(self, element: QtWidgets) -> None:
        """ Shows an Element

        Args:
            element (QtWidgets): Elemtent to be shown
                e.g.: QFrame, QPushButton, ...
        """
        element.show()

    def la_cancel_edit(self) -> None:
        """ Cancels the edit Process and hides the edit frame
        """
        global STATUS_MSG
        self.la_edit_frame_2.hide()
        STATUS_MSG = []

    def extract_all_ene_values(self):
        global ALL_DATA_PROBE
        if ALL_DATA_PROBE:
            
            for row in ALL_DATA_PROBE:
                try:
                    if row["Kennung"].split()[0] not in self._ene and isinstance(int(row["Kennung"].split()[1]), int):
                        self._ene.append(row["Kennung"].split()[0])
                except:
                    pass
            return self._ene
        else:
            ALL_DATA_PROBE = DATABASE_HELPER.get_all_probes()
            for row in ALL_DATA_PROBE:
                try:
                    if row["Kennung"].split()[0] not in self._ene and isinstance(int(row["Kennung"].split()[1]), int):
                        self._ene.append(row["Kennung"].split()[0])
                except:
                    pass
            return self._ene

    def extract_all_pnr_years(self):
        if ALL_DATA_PROJECT_NR:
            for sheet in ALL_DATA_PROJECT_NR:
                try:
                    for _, row in enumerate(ALL_DATA_PROJECT_NR[sheet]["Projekt-Nr."]):
                        if row.split("-")[0] not in self._pnr:
                            self._pnr.append(row.split("-")[0])
                except Exception as ex:
                    pass
            return self._pnr               

    def check_la_enable(self) -> None:
        if self.check_la_db_path.isChecked():
            self.laborauswertung_path.setEnabled(True)
            self.choose_laborauswertung_path_btn.setEnabled(True)

        else:
            self.laborauswertung_path.setEnabled(False)
            self.choose_laborauswertung_path_btn.setEnabled(False)



class Probe(QtWidgets.QMainWindow): 
    def __init__(self, parent=None):
        super(Probe, self).__init__(parent)
        uic.loadUi(r'./views/select_probe.ui', self)

        self.setWindowTitle(f"CapZa - Zasada - { __version__ } - Wähle Probe")

        self.load_probe_btn.clicked.connect(self.load_probe)
        self.cancel_btn.clicked.connect(self.close_window)
        init_shadow(self.load_probe_btn)
        init_shadow(self.cancel_btn)

    def init_data(self, dataset: list[dict]) -> None:
        """ Inputs all the Probe data into the TableWidget

        Args:
            dataset (list[dict]): List of dictionaries with Probevalues
        """

        self.model = QtGui.QStandardItemModel(len(dataset), 3)
        self.model.setHorizontalHeaderLabels(["Datum", "Kennung", "Materialbezeichnung"])
        for row, text in enumerate(dataset):
            date_item = QtGui.QStandardItem(text["Datum"])
            material_item = QtGui.QStandardItem(text["Materialbezeichnung"])
            kennung_item = QtGui.QStandardItem(text["Kennung"])
            self.model.setItem(row, 0, date_item)
            self.model.setItem(row, 2, material_item)
            self.model.setItem(row, 1, kennung_item)
        #filter proxy model
        self.filter_proxy_model = QtCore.QSortFilterProxyModel()
        self.filter_proxy_model.setSourceModel(self.model)
        self.filter_proxy_model.setFilterKeyColumn(1) # second column
        self.tableView.setModel(self.filter_proxy_model)
        self.results_lbl.setText(f"{self.filter_proxy_model.rowCount()} Ergebnisse gefunden")

        self.probe_filter_lineedit.textChanged.connect(self.apply_filter)

    def apply_filter(self, text):
        self.filter_proxy_model.setFilterRegExp(text)
        self.results_lbl.setText(f"{self.filter_proxy_model.rowCount()} Ergebnisse gefunden") 

    def load_probe(self) -> None:
        """ Gets the selected Probe and closes the Probe window
        """
        try:
            global SELECTED_PROBE
            index=(self.tableView.selectionModel().currentIndex())
            kennung=index.sibling(index.row(),1).data()
            material=index.sibling(index.row(),2).data()
            date=index.sibling(index.row(),0).data()
            SELECTED_PROBE= DATABASE_HELPER.get_specific_probe(id = kennung, material = material, date=date)

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
        global ALL_DATA_PROJECT_NR
        """ Loads the Nachweis data from ProjektNr.

        Args:
            projektnummer (str): Project Nr.
                e.g.: "22-0000"
        """ 


        global SELECTED_NACHWEIS
        for sheet in ALL_DATA_PROJECT_NR:
            try:
                if str(projektnummer) in ALL_DATA_PROJECT_NR[sheet]["Projekt-Nr."].values:
                    projekt_data = ALL_DATA_PROJECT_NR[sheet][ALL_DATA_PROJECT_NR[sheet]['Projekt-Nr.'] == str(projektnummer)]
            except Exception as ex:
                STATUS_MSG = [f"Probe mit Sheet {sheet} konnte nicht geladen werden: [{ex}]"]

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
        error_long_msg = "Es wurden mehrere Fehler gefunden: \n\n"
        if len(STATUS_MSG) > 1:
            for error in STATUS_MSG:
                error_long_msg+=f"- {str(error)}\n"
        elif len(STATUS_MSG) == 1:
            error_long_msg = str(STATUS_MSG[0])
        else: 
            error_long_msg = "/"

        self.error_lbl.setText(error_long_msg)
        init_shadow(self.close_error_info_btn)
        init_shadow(self.error_msg_frame)
        self.close_error_info_btn.clicked.connect(self.delete_error)

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

class Laborauswertung(QtWidgets.QDialog):
    def __init__(self, parent=None, la_type="add"):
        super(Laborauswertung, self).__init__(parent)
        uic.loadUi(r'./views/laborauswertung.ui', self)
        self.setWindowTitle(f"CapZa - Zasada - {__version__} - Laborauswertung")

        init_shadow(self.form_frame_1)
        init_shadow(self.form_frame_2)
        init_shadow(self.form_frame_3)

        self.save_la_btn.setEnabled(False)

        self.la_type = la_type

        self.shown = False
        self.show_calculation_frame_btn.clicked.connect(self.toggle_calculation_data)
        self.la_calculate_btn.clicked.connect(self.la_calculate)
        self.import_icp_scan_btn.clicked.connect(self.import_icp_scan)
        self.import_rfa_scan_btn.clicked.connect(self.import_rfa_scan)
        

        if self.la_type == "add":
            self.form_frame_3.show()
            self.la_aktion_lbl.setText("hinzufügen")
            self.save_la_btn.clicked.connect(self.la_add_save)

        else:
            self.la_aktion_lbl.setText("bearbeiten")
            self.save_la_btn.clicked.connect(self.la_edit_save)

    def toggle_calculation_data(self):
        if self.shown == True:
            self.form_frame_3.hide()
            self.shown = False
            self.show_calculation_frame_btn.setText("+")
        else:
            self.form_frame_3.show()
            self.shown = True
            self.show_calculation_frame_btn.setText("-")

    def la_calculate(self) -> None:
        """ Makes all calculations

        Args:
            item (__type__): Changed item in Laborauswertung Table
        """
        try:
            try:
                ### Berechnung % TS:
                if self.la_auswaage_fs_input.text() and self.la_einwaage_fs_input.text():
                    ts = float(self.la_auswaage_fs_input.text()) / (float(self.la_einwaage_fs_input.text())/100)
                    self.la_result_ts.setText(str(ts))
                    # Wasserfaktor
                    wasserfaktor = float(self.la_einwaage_fs_input.text()) / float(self.la_auswaage_fs_input.text())
                    self.la_result_wasserfaktor.setText(str(wasserfaktor))
            except Exception as ex:
                raise Exception(f"Fehler bei Berechung des TS und des Wasserfaktors: [{ex}]")
        
            try:
                ### Berechnung Wasserfaktor_getrocknet:
                if self.la_ts_der_probe_input.text():
                    wf_getrocknet = 100 / float(self.la_ts_der_probe_input.text())
                    self.la_result_wf_getrocknet.setText(str(wf_getrocknet))
            except Exception as ex:
                raise Exception(f"Fehler bei Berechung des Wasserfaktor getrocknetes Material: [{ex}]")

            try:
                ### Berechnung Lipos TS %: TODO: (Auswage-tara)/(Einwaage Soxlett/100)
                if self.la_lipos_auswaage_input.text() and self.la_lipos_tara_input.text() and self.la_einwaage_sox_input.text():
                    lipos_ts = (float(self.la_lipos_auswaage_input.text()) - (float(self.la_lipos_tara_input.text())))/(float(self.la_einwaage_sox_input.text()) / 100)
                    self.la_result_lipos_ts.setText(str(lipos_ts))
            except Exception as ex:
                raise Exception(f"Fehler bei Berechung des Lipos TS: [{ex}]")
            try:
                ### Berechnung GV %:
                if self.la_gv_auwaage_input.text() and self.la_gv_tara_input.text() and self.la_gv_einwaage_input.text():
                    gv = 100 - (float(self.la_gv_auwaage_input.text()) - float(self.la_gv_tara_input.text())) / float(self.la_gv_einwaage_input.text()) * 100
                    self.la_result_gv.setText(str(gv))
            except Exception as ex:
                raise Exception(f"Fehler bei Berechung des GVs: [{ex}]")

            
            try:
                ### Berechnung TDS:
                if self.la_tds_auswaage_input.text() and self.la_tds_tara_input.text() and self.la_tds_einwaage_input.text():
                    tds = (float(self.la_tds_auswaage_input.text()) - float(self.la_tds_tara_input.text())) * (1000 / float(self.la_tds_einwaage_input.text()) * 1000)
                    self.la_result_tds_gesamt.setText(str(tds))
            except Exception as ex:
                raise Exception(f"Fehler bei Berechung des DS Gesamt gelöste Stoffe: [{ex}]")

            try:
                ### Berechnung Einwaage TS
                if self.la_eluat_einwaage_os_input.text() and self.la_result_ts.text():
                    einwaage_ts = float(self.la_eluat_einwaage_os_input.text()) * (float(self.la_result_ts.text()) / 100)
                    self.la_result_einwaage_ts.setText(str(einwaage_ts))
            except Exception as ex:
                raise Exception(f"Fehler bei Berechung der Einwaage TS: [{ex}]")

            try:
                ### Berechnug Faktor
                if self.la_eluat_einwaage_os_input.text() and self.la_result_einwaage_ts.text():
                    faktor = float(self.la_eluat_einwaage_os_input.text()) / float(self.la_result_einwaage_ts.text())
                    self.la_result_faktor.setText(str(faktor))
            except Exception as ex:
                raise Exception(f"Fehler bei Berechung des Faktors: [{ex}]")
            
            self.save_la_btn.setEnabled(True)

        
        except Exception as ex:
            STATUS_MSG.append(ex)
            self.parent().feedback_message("error", STATUS_MSG)

    def insert_values(self, data: dict) -> None:
        """ Inserts all values when editing

        Args:
            data (dict): selected date, kennung and material
        """

        all_db_data = DATABASE_HELPER.get_specific_probe(id= data["kennung"], material=data["material"], date=data["datum"])
        y,m,d = all_db_data["Datum"].split()[0].split("-")
        self.la_date_edit.setDate(QDate(int(y),int(m),int(d)))
        try:
            self.la_material_input.setText(format_specific_insert_value("Materialbezeichnung", all_db_data))

            self.la_kennung_input.setText(format_specific_insert_value("Kennung", all_db_data))
            #---------------------------------------------#
            self.la_feuchte_input.setText(format_specific_insert_value("Wassergehalt %", all_db_data))
            self.la_fluorid_input.setText(format_specific_insert_value("Fluorid mg/L", all_db_data))
            
            self.la_ph_input.setText(format_specific_insert_value("pH-Wert", all_db_data))
            self.la_lf_input.setText(format_specific_insert_value("Leitfähigkeit mS/cm", all_db_data))
            self.la_cl_input.setText(format_specific_insert_value("Chlorid mg/L", all_db_data))
            self.la_cr_input.setText(format_specific_insert_value("Cr 205.560 (Aqueous-Axial-iFR)", all_db_data))
            self.la_doc_input.setText(format_specific_insert_value("Bezogen auf das eingewogene Material DOC mg/L", all_db_data))
            self.la_mo_input.setText(format_specific_insert_value("Mo 202.030 (Aqueous-Axial-iFR)", all_db_data))
            self.la_toc_input.setText(format_specific_insert_value("TOC %", all_db_data))
            self.la_ec_input.setText(format_specific_insert_value("EC %", all_db_data))
            #----------------------------------------------#
            # Berechnungs Daten
            self.la_ts_der_probe_input.setText(format_specific_insert_value(" TS der Probe", all_db_data))
            self.la_einwaage_sox_input.setText(format_specific_insert_value("Einwaage Soxlett", all_db_data))
            self.la_lipos_tara_input.setText(format_specific_insert_value("lipos_tara", all_db_data))
            self.la_lipos_auswaage_input.setText(format_specific_insert_value("lipos_auswaage", all_db_data))
            self.la_gv_tara_input.setText(format_specific_insert_value("GV Tara g", all_db_data))
            self.la_gv_einwaage_input.setText(format_specific_insert_value("GV Einwaage g", all_db_data))
            self.la_gv_auwaage_input.setText(format_specific_insert_value("GV Auswaage g", all_db_data))
            self.la_tds_tara_input.setText(format_specific_insert_value("TDS Tara g", all_db_data))
            self.la_tds_einwaage_input.setText(format_specific_insert_value("TDS Einwaage g", all_db_data))
            self.la_tds_auswaage_input.setText(format_specific_insert_value("TDS Auswaage g", all_db_data))
            self.la_einwaage_fs_input.setText(format_specific_insert_value("Einwaage FS g", all_db_data))
            self.la_auswaage_fs_input.setText(format_specific_insert_value("Auswaage FS g", all_db_data))
            self.la_eluat_einwaage_os_input.setText(format_specific_insert_value("Eluat Einwaage OS", all_db_data))
            #-------------------------------------------------#
            # Ergebnis Daten
            self.la_result_ts.setText(format_specific_insert_value("% TS", all_db_data))
            self.la_result_wasserfaktor.setText(format_specific_insert_value("Wasserfaktor", all_db_data))
            self.la_result_wf_getrocknet.setText(format_specific_insert_value("Wasserfaktor getrocknetes Material", all_db_data))
            self.la_result_lipos_ts.setText(format_specific_insert_value("Lipos TS %", all_db_data))
            self.la_result_gv.setText(format_specific_insert_value("GV %", all_db_data))
            self.la_result_tds_gesamt.setText(format_specific_insert_value("TDS Gesamt gelöste Stoffe mg/L", all_db_data))
            self.la_result_einwaage_ts.setText(format_specific_insert_value("Einwaage TS", all_db_data))
            self.la_result_faktor.setText(format_specific_insert_value("Faktor", all_db_data))
            #-----------------------------------------------#
            # Bemerkungen
            bemerkung_lst = all_db_data[r"strukt_bemerkung"].split(";")
            self.listWidget.addItems(bemerkung_lst) if all_db_data[r"strukt_bemerkung"] != None else "-"
            
        except Exception as ex:
            print(ex)

    def la_collect_inserted_values(self) -> None:
        """ Adds the new Laborauswertung entry to the database
        """

        global STATUS_MSG, ALL_DATA_PROBE

        STATUS_MSG = []
        try:
            datum = self.la_date_edit.date().toString("yyyy-MM-dd 00:00:00")
            material = self.la_material_input.text()
            kennung = self.la_kennung_input.text()
            #---------------------------------------------#
            feuchte = self.la_feuchte_input.text()
            fluorid = self.la_fluorid_input.text()
            ph = self.la_ph_input.text()
            leitf = self.la_lf_input.text()
            chlorid = self.la_cl_input.text()
            cr = self.la_cr_input.text()
            doc = self.la_doc_input.text()
            mo = self.la_mo_input.text()
            toc = self.la_toc_input.text()
            ec = self.la_ec_input.text()
            #----------------------------------------------#
            # Berechnungs Daten
            ts_probe = self.la_ts_der_probe_input.text()
            einwaage_sox = self.la_einwaage_sox_input.text()
            lipos_tara = self.la_lipos_tara_input.text()
            lipos_auswaage = self.la_lipos_auswaage_input.text()
            gv_tara = self.la_gv_tara_input.text()
            gv_einwaage = self.la_gv_einwaage_input.text()
            gv_auswaage = self.la_gv_auwaage_input.text()
            tds_tara = self.la_tds_tara_input.text()
            tds_einwaage = self.la_tds_einwaage_input.text()
            tds_auswaage = self.la_tds_auswaage_input.text()
            einwaage_fs = self.la_einwaage_fs_input.text()
            auswaage_fs = self.la_auswaage_fs_input.text()
            eluat_einwaage = self.la_eluat_einwaage_os_input.text()
            #-------------------------------------------------#
            # Ergebnis Daten
            result_ts = self.la_result_ts.text()
            result_wasserfaktor = self.la_result_wasserfaktor.text()
            result_wasserfaktor_getrocknet = self.la_result_wf_getrocknet.text()
            result_lipos_ts = self.la_result_lipos_ts.text()
            result_gv = self.la_result_gv.text()
            result_tds_gesamt = self.la_result_tds_gesamt.text()
            result_einwaage_ts = self.la_result_einwaage_ts.text()
            result_faktor = self.la_result_faktor.text()
            
            # Bemerkung Daten
            bemerkung_hist = [self.listWidget.item(i).text() for i in range(self.listWidget.count())]
            if self.bemerkung_input.text() != None:
                if len(bemerkung_hist) > 0:
                    bemerkung = ";".join(bemerkung for bemerkung in bemerkung_hist) + f";[{TODAY_FORMAT_STRING}]: {self.bemerkung_input.text()}"
                else:
                    bemerkung = f"[{TODAY_FORMAT_STRING}]: {self.bemerkung_input.text()}"
            else: bemerkung = ""
            self.save_data = {
                "Datum": datum,
                "Materialbezeichnung": material,
                "Kennung": kennung,
                #-------------
                "Wassergehalt %": feuchte,
                "Fluorid mg/L": fluorid,
                "pH-Wert": ph,
                "Leitfähigkeit mS/cm": leitf,
                "Chlorid mg/L": chlorid,
                "Cr 205.560 (Aqueous-Axial-iFR)": cr,
                "Bezogen auf das eingewogene Material DOC mg/L": doc,
                "Mo 202.030 (Aqueous-Axial-iFR)": mo,
                "TOC %": toc,
                "EC %": ec,
                #-------------
                " TS der Probe":ts_probe,
                "Einwaage Soxlett":einwaage_sox,
                "lipos_tara":lipos_tara,
                "lipos_auswaage":lipos_auswaage,
                "GV Tara g":gv_tara,
                "GV Einwaage g":gv_einwaage,
                "GV Auswaage g":gv_auswaage,
                "TDS Tara g": tds_tara,
                "TDS Einwaage g":tds_einwaage,
                "TDS Auswaage g":tds_auswaage,
                "Einwaage FS g":einwaage_fs,
                "Auswaage FS g":auswaage_fs,
                "Eluat Einwaage OS":eluat_einwaage,
                #--------------------------------#
                r"% TS":result_ts,
                r"Wasserfaktor":result_wasserfaktor,
                r"Wasserfaktor getrocknetes Material":result_wasserfaktor_getrocknet,
                r"Lipos TS %":result_lipos_ts,
                r"GV %":result_gv,
                r"TDS Gesamt gelöste Stoffe mg/L":result_tds_gesamt,
                r"Einwaage TS":result_einwaage_ts,
                r"Faktor":result_faktor,
                #----------------------------------#
                r"strukt_bemerkung": bemerkung
            }                
        except Exception as ex:
            STATUS_MSG.append(f"Fehler beim Auslesen der Daten: [{ex}]")
            self.parent()._check_for_errors()
            return

    def la_edit_save(self) -> None:
        self.la_collect_inserted_values()
        DATABASE_HELPER.edit_laborauswertung(self.save_data, self.save_data["Kennung"], self.save_data["Datum"])
        self.close()

    def la_add_save(self) -> None:
        self.la_collect_inserted_values()
        DATABASE_HELPER.add_laborauswertung(self.save_data)
        self.close()

    def close_window(self) -> None:
        self.close()

    def import_icp_scan(self) -> str:
        file = QFileDialog.getOpenFileName(self, "ICP-Scan", "C://", "CSV Files (*.csv)")
        return file
    
    def import_rfa_scan(self) -> str:
        file = QFileDialog.getOpenFileName(self, "RFA-Scan", "C://", "CSV Files (*.csv)")
        return file

def format_specific_insert_value(col_name: str, all_data: dict) -> str:
    if type(all_data) == pd.DataFrame:
        all_data = all_data.to_dict(orient='records')[0]
    return str(all_data[col_name]) if all_data[col_name] != None else "-"

def round_if_psbl(value) -> str:
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
        elif value == "<LoD":
            return "-"
        else:
            return str(value)

def init_shadow(widget) -> None:
        """Sets shadow to the given widget

        Args:
            widget (QWidget): QFrame, QButton, ....
        """

        effect = QGraphicsDropShadowEffect()
        effect.setOffset(0, 1)
        effect.setBlurRadius(8)
        widget.setGraphicsEffect(effect)



if __name__ == "__main__":
    STATUS_MSG = []
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
            DATABASE_HELPER = DatabaseHelper(DB_PATH)
        except Exception as ex:
            STATUS_MSG.append(f"Die Datenbank konnte nicht gefunden werden. Bitte überprüfe in der Referenzeinstellung: [{ex}]")

        win = Ui()

        try:
            ALL_DATA_PROBE = DATABASE_HELPER.get_all_probes()
        except Exception as ex:
            print(ex)
            STATUS_MSG.append(f"Es konnten keine Proben geladen werden: [{ex}]")

        try:
            ALL_DATA_NACHWEIS = pd.read_excel(NW_PATH)
        except Exception as ex:
            print(ex)
            STATUS_MSG.append(f"Es wurde keine Nachweisliste gefunden. Bitte prüfe in den Referenzeinstellungen. [{str(ex)}]")

        if STATUS_MSG != []:
            win.feedback_message("error", STATUS_MSG)

        
        win._check_for_errors()
        
        splash.finish(win)
        win.show()
        sys.exit(app.exec_())
    else: pass