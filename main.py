
from modules.data_functions import format_specific_insert_value, round_if_psbl
from modules.ramses_helper import RamsesHelper
from modules.db_helper.__init__ import DatabaseHelper
from modules.config_helper import ConfigHelper
from modules.word_helper import Word_Helper
from modules.icp_scan import get_icp_data
from modules.read_rfa import get_rfa_data
from modules.update_helper import UpdateHelper


from datenklassen import projectnr
import logging

import datetime
from multiprocessing import Process, freeze_support
from docx2pdf import convert
import pandas as pd
from PyQt5 import QtWidgets, uic, QtCore, QtGui
from PyQt5.QtGui import QIcon, QPixmap, QFont, QStandardItemModel, QStandardItem, QIntValidator
from PyQt5.QtCore import Qt, QDate, QTimer, QThread
from PyQt5.QtWidgets import QFileDialog, QGraphicsDropShadowEffect, QSplashScreen, QDateEdit, QHeaderView, QComboBox, QPushButton
import os
import sys
import re

from datenklassen import laborauswertung

import assets.icons
import resources

from threading import Thread
import os
dirname = os.path.dirname(__file__)

# import helper modules
ROOT_DIR = os.path.dirname(os.path.abspath(__file__))

today_date_raw = datetime.datetime.now()
TODAY_FORMAT_STRING = today_date_raw.strftime(r"%d.%m.%Y")

LOCKFILE_PROBE = QtCore.QLockFile(QtCore.QDir.tempPath() + 'capza_probe.lock')
LOCKFILE_ERROR = QtCore.QLockFile(QtCore.QDir.tempPath() + 'capza_error.lock')


CONFIG_HELPER = ConfigHelper()
DATABASE_HELPER = None

RAMSES_CONN = None

__version__ = CONFIG_HELPER.get_specific_config_value("version")
__update__ = False

SELECTED_PROBE = None
SELECTED_NACHWEIS = None

ALL_DATA_PROBE = None
ALL_DATA_NACHWEIS = None
ALL_DATA_PROJECT_NR = None

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
        super(Ui, self).__init__(parent)
        uic.loadUi(r'./views/main.ui', self)

        self.init_main()

    def check_for_update(self):
        uh = UpdateHelper(ROOT_DIR, __version__)
        # delete Updater
        try:
            uh.delete_installer()
        except Exception as ex:
            print(ex)
        try:
            uh.is_new_version()
            if uh.is_new_version():
                print()
                dlg = QtWidgets.QMessageBox(self)
                dlg.setWindowTitle("Update verfügbar!")
                print("1", dlg)
                dlg.setText(
                    f"Eine neue Version {uh.get_new_version()} ist verfügbar!\n\n"
                    "Soll das Update jetzt heruntergeladen werden?")
                print("2", uh.get_new_version())
                dlg.setStandardButtons(QtWidgets.QMessageBox.Yes |
                                    QtWidgets.QMessageBox.No)
                print("3", dlg)
                dlg.setIcon(QtWidgets.QMessageBox.Question)
                print("4", dlg)
                button = dlg.exec()
                print("5", dlg)

                if button == QtWidgets.QMessageBox.Yes:
                    uh.clean_up()
                    uh.update()
                    # Process(target=uh.update).start()
                    return True
                else:
                    return False
        except Exception as ex:
            print("Update konnte nicht geprüft werden:", ex)


    def init_main(self) -> None:
        # STANDARDEINSTELLUNGEN:
        self.word_helper = Word_Helper()
        self.ramses_helper = RamsesHelper()

        self.setWindowTitle(f"CapZa - Zasada - { __version__ } ")
        self.setWindowIcon(QIcon(r'./assets/capza_icon.png'))
        self.stackedWidget.setCurrentIndex(0)
        self.showMaximized()

        self.logo_right_lbl.setPixmap(QPixmap("./assets/l_logo.png"))
        self.second_info_lbl.hide()
        self.error_info_btn.clicked.connect(self.showError)
        self.status_msg_btm.hide()
        self.hide_admin_msg_btn.clicked.connect(self.hide_admin_msg)
        self._ene = list()
        self._pnr = list()

        # NAVIGATION:
        self.nav_data_btn.clicked.connect(lambda: self.display(0))
        self.nav_analysis_btn.clicked.connect(lambda: self.display(1))
        self.nav_pnp_entry_btn.clicked.connect(lambda: self.display(2))
        self.nav_pnp_output_btn.clicked.connect(lambda: self.display(3))
        self.nav_order_form_btn.clicked.connect(lambda: self.display(4))
        self.nav_settings_btn.clicked.connect(lambda: self.display(5))
        self.nav_laborauswertung_btn.clicked.connect(lambda: self.display(6))

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

        # DATENEINGABE:
        self.end_dateedit.setDate(self._get_todays_date_qdate())
        self.load_project_nr()
        self.select_probe_btn.clicked.connect(self.read_all_probes)
        self.select_nachweis_btn.clicked.connect(self.read_all_nachweis)
        self.brandtest_combo.currentTextChanged.connect(
            lambda: self.analysis_brandtest_lineedit.setText(self.brandtest_combo.currentText()))
        self.nh3_lineedit_2.textChanged.connect(self.check_nh3_value)
        self.h2_lineedit_2.textChanged.connect(
            lambda: self.h2_lineedit.setText(self.h2_lineedit_2.text()))
        self.kennung_rb.clicked.connect(lambda: self.empty_manual_search(
            self.pnr_combo, self.project_nr_lineedit))
        self.pnr_rb.clicked.connect(lambda: self.empty_manual_search(
            self.kennung_combo, self.kennung_lineedit))

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

        # ANALYSEWERTE:
        self.migrate_btn.clicked.connect(self.create_bericht_document)
        self.aqs_btn.clicked.connect(self.create_aqs_document)

        # PNP Input:
        self.pnp_in_create_protokoll.clicked.connect(self.trigger_pnp_in)

        # PNP Output:
        int_validator = QIntValidator(0, 999999999, self)
        self.output_nr_lineedit.setValidator(int_validator)
        self.pnp_out_protokoll_btn.clicked.connect(
            self.create_pnp_out_protokoll)

        # AUFTRAGSFORMULAR:
        self.autrag_load_column_view()
        self.auftrag_add_auftrag_btn.clicked.connect(self.auftrag_add_auftrag)

        # LABORAUSWERTUNG:
        self.la_table_view.doubleClicked.connect(self.edit_laborauswertung)
        self.add_laborauswertung_btn.clicked.connect(self.add_laborauswertung)
        self.la_table_view.horizontalHeader().setStyleSheet(
            "QHeaderView {font-size: 12pt; font-weight:bold} QTableView {border: 1px solid black;}")
        self.la_table_view.resizeColumnsToContents()
        init_shadow(self.la_table_view)
        self.la_changed_item_lst = {}

        # REFERENZEINSTELLUNGEN:
        self.project_nr_path.setText(PNR_PATH)
        self.save_bericht_path.setText(STANDARD_SAVE_PATH)
        self.laborauswertung_path.setText(LA_PATH)
        self.db_path.setText(DB_PATH)
        self.disable_settings_lines()

        self.choose_project_nr_btn.clicked.connect(self.choose_project_nr)
        self.choose_laborauswertung_path_btn.clicked.connect(self.choose_la)
        self.choose_db_path_btn.clicked.connect(self.choose_db)
        self.choose_save_bericht_path.clicked.connect(lambda: self.select_folder(
            self.save_bericht_path, "Wähle den Standardpfad zum Speichern aus."))

        self.clear_cache_btn.clicked.connect(self._clear_cache)
        self.save_references_btn.clicked.connect(self.save_references)

        self.check_la_enable()
        self.check_la_db_path.toggled.connect(self.check_la_enable)

        if self.project_nr_path.text() == "":
            STATUS_MSG.append(
                "Es ist keine Projektliste hinterlegt. Prüfe in den Referenzeinstellungen.")
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
        global SELECTED_PROBE, SELECTED_NACHWEIS, ALL_DATA_PROBE, ALL_DATA_NACHWEIS, ALL_DATA_PROJECT_NR, PNR_PATH, STATUS_MSG, BERICHT_FILE, ALIVE, PROGRESS
        SELECTED_PROBE = 0
        SELECTED_NACHWEIS = 0
        PNR_PATH = ""
        STATUS_MSG = []
        BERICHT_FILE = ""

    def _get_todays_date_qdate(self) -> QDate:
        """ Get QDate object from current date string

        Returns:
            QDate: Contains the current Date
        """

        d, m, y = TODAY_FORMAT_STRING.split(".")
        return QDate(int(y), int(m), int(d))

    def trigger_pnp_in(self) -> bool:
        """ Opens a dialog after setting the pnp in-
        """
        dlg = QtWidgets.QMessageBox(self)
        dlg.setWindowTitle("PNP Input mit oder ohne Weiterberechnungsformular")
        if self.weiterberechnung_checkBox.checkState() == 2:
            dlg.setText(
                "Soll das Weiterberechnungsformular wirklich dazu erstellt werden?")
        else:
            dlg.setText(
                "Soll wirklich kein Weiterberechnungsformular dazu erstellt werden? [Wenn nicht, dann wähle 'No']s")
        dlg.setStandardButtons(QtWidgets.QMessageBox.Yes |
                               QtWidgets.QMessageBox.No)
        dlg.setIcon(QtWidgets.QMessageBox.Question)
        button = dlg.exec()

        if button == QtWidgets.QMessageBox.Yes:
            self.create_weiterberechnung_document()
            self.create_pnp_in_protokoll()
            return True
        else:
            self.create_pnp_in_protokoll()
            return False

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
                    SELECTED_PROBE = DATABASE_HELPER.get_specific_probe(
                        kennung)
                    nachweis_data = ALL_DATA_NACHWEIS[ALL_DATA_NACHWEIS['en_nachweisnummer'] == get_full_project_ene_number(kennung)[0]]
                    nachweis_data['nachweis_letters'] = kennung_letters
                    nachweis_data['nachweis_nr'] = kennung_nr
                    nachweis_data['project_yr'] = "-"
                    nachweis_data['project_nr'] = "-"

                    SELECTED_NACHWEIS = nachweis_data
                    SELECTED_NACHWEIS["menge"] = str((float(SELECTED_NACHWEIS["ve_anfall_kjahr1"])+float(SELECTED_NACHWEIS["ve_anfall_kjahr2"])+float(
                        SELECTED_NACHWEIS["ve_anfall_kjahr3"])+float(SELECTED_NACHWEIS["ve_anfall_kjahr4"])+float(SELECTED_NACHWEIS["ve_anfall_kjahr5"]))/5)
                    self.insert_values(kennung)
                else:
                    mark_error_line(self.kennung_combo, "QComboBox")
                    raise Exception("Wählen Sie die Kennungsart aus")
            elif self.pnr_rb.isChecked():
                if project_year != "-":
                    SELECTED_PROBE = DATABASE_HELPER.get_specific_probe(
                        projectnr)
                    print(SELECTED_PROBE)
                    logging.info(SELECTED_PROBE)
                    for sheet in ALL_DATA_PROJECT_NR:
                        try:
                            if str(projectnr) in ALL_DATA_PROJECT_NR[sheet]["Projekt-Nr."].values:
                                nachweis_data = ALL_DATA_PROJECT_NR[sheet][ALL_DATA_PROJECT_NR[sheet]['Projekt-Nr.'] == str(
                                    projectnr)]
                        except Exception as ex:
                            STATUS_MSG = [
                                f"Probe mit Sheet {sheet} konnte nicht geladen werden: [{ex}]"]
                            self.feedback_message("error", STATUS_MSG)
                    nachweis_data["en_ort"] = nachweis_data["Ort"]
                    nachweis_data["en_plz"] = ""
                    nachweis_data["en_name"] = nachweis_data['Erzeuger']
                    nachweis_data["en_asn_info_feld"] = nachweis_data['AVV']
                    nachweis_data["ve_abfallbez_betr_intern"] = nachweis_data['Material']
                    nachweis_data['nachweis_letters'] = "-"
                    nachweis_data['nachweis_nr'] = "-"
                    nachweis_data['project_yr'] = project_year
                    nachweis_data['project_nr'] = project_nr
                    SELECTED_NACHWEIS = nachweis_data
                    SELECTED_NACHWEIS["menge"] = str(SELECTED_NACHWEIS["Menge [t/a]"].values[0])
                    self.insert_values(projectnr, "projekt")
                else:
                    mark_error_line(self.pnr_combo, "QComboBox")
                    raise Exception("Wählen Sie das Projektjahr aus")
            else:
                raise Exception(
                    "Es wurde keine Suchart ausgewählt. Wählen Sie eine Suchart aus.")
        except Exception as ex:
            STATUS_MSG = [f"Fehler: {str(ex)}"]
            self.feedback_message("error", STATUS_MSG)
            self.empty_values()

    def empty_manual_search(self, widget_combo, widget_line):
        # empty values
        widget_combo.setCurrentText("-")
        widget_line.setText("")

    def get_full_project_ene_number(self, nummer: str) -> str:
        """ Gets the whole , VNE, ... Projectnr. from the shortform """

    def _no_function(self) -> None:
        """ Mock function for features, that are not yet implemented
        """

        global STATUS_MSG
        STATUS_MSG.append("Diese Funktion steht noch nicht zu verfügung.")
        self.feedback_message("attention", STATUS_MSG)

    def choose_la(self) -> None:
        """ Choose LA_PATH from Referenzeinstellungen
        """

        self.select_file(self.laborauswertung_path, "",
                         "Wähle die Laborauswertung aus...", "Excel Files (*.xlsx *.xls)")

    def choose_db(self) -> None:
        """ Choose DB_PATH from Referenzeinstellungen
        """

        global DB_PATH
        DB_PATH = self.select_file(
            self.db_path, "", "Wähle die Datenbank aus...", "Databse Files (*.db)")

    def choose_project_nr(self) -> None:
        """ Choose PNR_PATH from Referenzeinstellungen
        """

        global PNR_PATH
        PNR_PATH = self.select_file(self.project_nr_path, "", "Wähle die Projektnummernliste aus...", "Excel Files (*.xlsx *.xls)")


    def load_project_nr(self) -> None:
        """ Loads data from Projekt.xlsx to CapZa
        """

        global ALL_DATA_PROJECT_NR
        try:
            ALL_DATA_PROJECT_NR = pd.read_excel(PNR_PATH, sheet_name=None)
        except Exception as ex:
            STATUS_MSG.append(
                f"Projektnummern konnten nicht geladen werden: [{ex}]")
            self.feedback_message("error", STATUS_MSG)

    def showError(self) -> None:
        """ Shows the error frame
        """
        if LOCKFILE_ERROR.tryLock(100):
            error = Error(self)
            error.show()
        else:
            pass

    def check_for_errors(self) -> None:
        """ Checks for possible errors and schows them in case there are any
        """

        global STATUS_MSG
        if len(STATUS_MSG) > 0:
            self.error_info_btn.show()
        else:
            self.error_info_btn.hide()

    def disable_settings_lines(self) -> None:
        """ Disables Line edits from Settings 
        """

        self.project_nr_path.setEnabled(False)

    def select_folder(self, line, title: str) -> None:
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
        """ Takes all the text from the Referenzsettings and saves it to the capza_config.ini """

        global DATABASE_HELPER
        global STATUS_MSG, ALL_DATA_PROBE
        save_path = ""
        project_nr_path = ""
        la_path = ""
        db_path = ""

        try:
            if self.project_nr_path.text():
                project_nr_path = self.project_nr_path.text()
                PNR_PATH = project_nr_path
            if self.save_bericht_path.text():
                save_path = self.save_bericht_path.text()
                STANDARD_SAVE_PATH = save_path
            if self.laborauswertung_path.text() and self.check_la_db_path.isChecked():
                la_path = self.laborauswertung_path.text()
                ALL_DATA_PROBE = DATABASE_HELPER.excel_to_sql(la_path)
                LA_PATH = la_path
            if self.db_path.text():
                db_path = self.db_path.text()
                DB_PATH = db_path
                DATABASE_HELPER = DatabaseHelper(DB_PATH)
                ALL_DATA_PROBE = DATABASE_HELPER.get_all_probes()

            references = {
                "project_nr_path": project_nr_path,
                "save_path": save_path,
                "la_path": la_path,
                "db_path": db_path
            }

            for key, value in references.items():
                CONFIG_HELPER.update_specific_value(key, value)

            self.laborauswertung_path.setText("")
            STATUS_MSG = []

            t_1 = Thread(target=self.init_main)
            t_1.start()
            t_2 = Thread(target=self.feedback_message, args=("success", ["Neue Referenzen erfolgreich gespeichert."],))
            t_2.start()

        except Exception as ex:
            STATUS_MSG.append("Das Speichern ist fehlgeschlagen: " + str(ex))
            self.feedback_message("error", f"Fehler beim Speichern: [{ex}]")

    def open_pop_up_probe(self, dataset: list[dict]) -> None:
        """ Opens the Probe window with the entire dataset

        Args:
            dataset (dict): Dataset from the database
        """

        global STATUS_MSG

        try:
            if LOCKFILE_PROBE.tryLock(100):
                self.probe_pop_up = PopUp(self)
                self.probe_pop_up.show()
                self.probe_pop_up.init_probe_data(dataset)
            else:
                pass
        except Exception as ex:
            STATUS_MSG.append(
                f"Es  konnten keine Daten gefunden werden. Importiere ggf. eine Laborauswertungsexcel: [{ex}]")
            self.feedback_message("error", STATUS_MSG)

    def open_pop_up_nachweis(self, dataset: pd.DataFrame) -> None:
        """ Opens the Probe window with the entire dataset

        Args:
            dataset (pd.DataFrame): Dataset from the database
        """

        global STATUS_MSG

        try:
            if LOCKFILE_PROBE.tryLock(100):
                self.nachweis_pop_up = PopUp(self, "Nachweis")
                self.nachweis_pop_up.show()
                self.nachweis_pop_up.init_nachweis_data(dataset)

            else:
                pass
        except Exception as ex:
            STATUS_MSG.append(
                f"Es  konnten keine Daten gefunden werden. Importiere ggf. eine Laborauswertungsexcel: [{ex}]")
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

        # if i == 5:
        #     self.show_second_info(
        #         "Der Pfad zur 'Nachweis Übersicht' Excel ist nur temporär und wird in Zukunft durch Echtdaten aus RAMSES ersetzt.")

        if i == 6:
            t_1 = Thread(target=self.load_laborauswertung)
            t_2 = Thread(target=self.feedback_message, args=(
                "info", ["Laborauswertung wird geladen..."],))
            t_1.start()
            t_2.start()

    def create_bericht_document(self) -> None:
        """ Builds and creates the Bericht file. Therefore it gathers all data from the FE.
        """

        global STATUS_MSG
        id = "x" if self.id_check_2.isChecked() else ""
        vorpruefung = "x" if self.precheck_check_2.isChecked() else ""

        ahv = "x" if self.ahv_check_2.isChecked() else ""
        erzeuger = "x" if self.erzeuger_check_2.isChecked() else ""

        nh3 = str(self.nh3_lineedit.text())
        h2 = str(self.h2_lineedit.text())
        brandtest = str(self.brandtest_combo.currentText())
        farbe = str(self.color_lineedit.text())
        konsistenz = str(self.consistency_lineedit.text())
        geruch = str(self.smell_lineedit.text())
        bemerkung = str(self.remark_textedit.toPlainText())

        aoc = 0
        toc = 0
        ec = 0
        if (not self.toc_lineedit.text() == "-") and (not self.toc_lineedit.text() == None):
            toc = round_if_psbl(float(self.toc_lineedit.text()))
        else:
            toc = ""

        if (not self.ec_lineedit.text() == "-") and (not self.ec_lineedit.text() == None):
            print(not self.ec_lineedit.text() == "-")
            print(not self.ec_lineedit.text() == None)
            print(self.ec_lineedit.text())
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

        date_str = str(SELECTED_PROBE["datum"])
        format = r"%Y-%m-%d %H:%M:%S"
        date_dt = datetime.datetime.strptime(date_str, format)
        date = datetime.datetime.strftime(date_dt, r"%d.%m.%Y")

        data = {
            "projekt_nr": str(SELECTED_PROBE["material_kenn"]),
            "bezeichnung": self.name_lineedit.text(),
            "erzeuger_name": str(list(SELECTED_NACHWEIS["en_name"])[0]),
            #
            "id": id,
            "vorpruefung": vorpruefung,
            "ahv": ahv,
            "erzeuger": erzeuger,
            "avv": self.format_avv_space_after_every_second(str(SELECTED_NACHWEIS["en_asn_info_feld"].values[0])),
            "menge": str(SELECTED_NACHWEIS["menge"].values[0]),
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

        word_file = self.create_word(CONFIG_HELPER.get_specific_config_value("bericht_vorlage"), data, "Bericht")
        if word_file != "/":
            try:
                thread1 = Thread(
                    target=self.word_helper.open_word, args=(word_file,))
                thread2 = Thread(target=self.feedback_message,
                                 args=("info", ["Bericht wird geöffnet..."]))
                thread1.start()
                thread2.start()
            except Exception as ex:
                self.feedback_message(
                    "attention", [f"Fehler beim Erstellen der Word Datei [{ex}]"])
                STATUS_MSG.append(str(ex))

    def create_aqs_document(self) -> None:
        """ Builds and creates the AQS Bericht. Therefore gathers all data from FE
        """

        global STATUS_MSG

        pnp_check = "Ja" if self.pnp_check.checkState() == 2 else "Nein"

        data = {
            "projekt_nr": str(SELECTED_PROBE["material_kenn"]),
            "bezeichnung": self.name_lineedit.text(),
            "datum": TODAY_FORMAT_STRING,
            "probe_vorhanden ": pnp_check,
            "income_datum":self.probe_date.date().toString("dd.MM.yyyy"),
            "avv_code": self.format_avv_space_after_every_second(str(SELECTED_NACHWEIS["en_asn_info_feld"].values[0])),
            "erzeuger": str(list(SELECTED_NACHWEIS["en_name"])[0]),
            "tonnage": str(SELECTED_NACHWEIS["menge"].values[0])
        }
        word_file = self.create_word(CONFIG_HELPER.get_specific_config_value("aqs_vorlage"), data, "AQS")
        if word_file != "/":
            try:
                thread1 = Thread(
                    target=self.word_helper.open_word, args=(word_file,))
                thread2 = Thread(target=self.feedback_message,
                                 args=("info", ["AQS wird geöffnet..."]))
                thread1.start()
                thread2.start()
            except Exception as ex:
                self.feedback_message(
                    "attention", [f"Fehler beim Erstellen der Word Datei [{ex}]"])
                STATUS_MSG.append(str(ex))

    def autrag_load_column_view(self) -> None:
        """ Loads the Column View
        """

        self.model = QStandardItemModel()
        self.model.setHorizontalHeaderLabels(
            ['Projekt-/Nachweisnummer(n)', 'Probenahmedatum', 'Analyseauswahl', 'ggf. spezifische Probenbezeichnung', 'Info 2', '#'])
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

        self.auftrag_table_view.setIndexWidget(
            self.model.index(row, 1), _date_edit)
        self.auftrag_table_view.setIndexWidget(
            self.model.index(row, 2), _combo1)
        self.auftrag_table_view.setIndexWidget(
            self.model.index(row, 4), _combo2)
        self.auftrag_table_view.setIndexWidget(
            self.model.index(row, 5), _delete_btn)

    def create_pnp_out_protokoll(self) -> None:
        """ Builds and creates the PNP-Output-Protocol. Therefore gathers all the data from the FE
        """

        anzahl = self.amount_analysis.currentText()
        vorlage_document = self._specific_vorlage(anzahl)
        # get all data

        # get Probenehmer
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
        # get Probenehmer
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
        word_file = self.create_word(vorlage_document, data, "PNP Output Protokoll")
        try:
            thread1 = Thread(
                target=self.word_helper.open_word, args=(word_file,))
            thread2 = Thread(target=self.feedback_message,
                             args=("info", ["PNP Output Protokoll wird geöffnet..."]))
            thread1.start()
            thread2.start()
        except Exception as ex:
            self.feedback_message(
                "attention", [f"Fehler beim Erstellen der Word Datei [{ex}]"])
            STATUS_MSG.append(str(ex))

    def create_pnp_in_protokoll(self) -> None:
        """ Builds and creates the PNP-Output-Protocol. Therefore gathers all the data from the FE
        """
        # GRUND
        grund_list = []
        if self.jahreskontrolle_check.isChecked():
            grund_list.append("Jahreskontrolle")
        if self.deklarationsanalyse_check.isChecked():
            grund_list.append("Deklarationsanalyse")
        if self.sonstige_grund_check.isChecked():
            grund_list.append(self.sonstige_grund_lineedit.text())
        

        probe_date = self.pnp_input_date_edit.date().toString("dd.MM.yyyy")
        ####
        # get Probenehmer
        if self.probenehmer_ms_pnp_in.isChecked():
            probenehmer = "M. Segieth"
        elif self.probenehmer_sg_pnp_in.isChecked():
            probenehmer = "S. Goritz"
        elif self.probenehmer_lz_pnp_in.isChecked():
            probenehmer = "L. Zasada"
        elif self.probenehmer_sonst_pnp_in.isChecked():
            probenehmer = self.sonst_pn_pnp_in_lineedit.text()
        else:
            probenehmer = "-"

        # get Probenehmer
        if self.anw_person_ms_pnp_in.isChecked():
            anwesende_personen = "M. Segieth"
        elif self.anw_person_sg_pnp_in.isChecked():
            anwesende_personen = "S. Goritz"
        elif self.anw_person_lz_pnp_in.isChecked():
            anwesende_personen = "L. Zasada"
        elif self.anw_person_sonst_pnp_in.isChecked():
            anwesende_personen = self.sonst_ap_pnp_in_lineedit.text()
        else:
            anwesende_personen = "-"

        probe_date = self.pnp_output_probenahmedatum.date().toString("dd.MM.yyyy")
        erzeuger_mit_stadt = self.person_lineedit.text() + ", " + self.location_lineedit.text()

        # Lagerung
        lagerung_list = []
        if self.haufwerk_check.isChecked():
            lagerung_list.append("Haufwerk")
        if self.bigbags_check.isChecked():
            lagerung_list.append("Big Bags")
        if self.silo_check.isChecked():
            lagerung_list.append("Silo")

        # Probenahmegerät
        probegeraet_list = []
        if self.schaufel_check.isChecked():
            probegeraet_list.append("Schaufel")
        if self.eimer_check.isChecked():
            probegeraet_list.append("Eimer")
        if self.bagger_check.isChecked():
            probegeraet_list.append("Bagger")
        if self.gerat_sonstig_check.isChecked():
            probegeraet_list.append(self.sontiges_gerat_le.text())

        data = {
            "grund":', '.join(grund_list),
            "datum": str(TODAY_FORMAT_STRING),
            "probennehmer": probenehmer,
            "anwesende_personen": anwesende_personen,
            "erzeuger_mit_stadt": erzeuger_mit_stadt,
            "nachweisnummer": SELECTED_NACHWEIS["en_nachweisnummer"].values[0],
            "avv_material": self.name_lineedit.text(),
            "farbe": self.pnp_in_color_le.text(),
            "geruch": self.pnp_in_smell_le.text(),
            "konsistenz": self.pnp_in_consistency_le.text(),
            "groesstkorn": self.pnp_in_groesstkor_le.text(),
            "lagerung": f"{self.pnp_in_menge_le.text()}, {', '.join(lagerung_list)}",
            "probenahmegerät": ', '.join(probegeraet_list),
            "probenvorbereitung": self.pnp_in_pnprep_le.text(),
            "probe_datum": probe_date,
            "bemerkung": self.pnp_in_bemerkung_te.toPlainText(),
            "entnahmetiefe": self.pnp_in_entnahmetiefe_le.text(),
            "pnp_in_probe": self.pnp_in_probe_kind_le.text()
        }
        word_file = self.create_word(CONFIG_HELPER.get_specific_config_value("pnp_in_vorlage"), data, "PNP_Input_Protokoll")
        try:
            thread1 = Thread(target=self.word_helper.open_word, args=(word_file,))
            thread2 = Thread(target=self.feedback_message, args=("info", ["PNP Eingang Protokoll wird geöffnet..."]))
            
            thread1.start()
            thread2.start()
        except Exception as ex:
            self.feedback_message(
                "attention", [f"Fehler beim Erstellen der Word Datei [{ex}]"])
            STATUS_MSG.append(str(ex))

    def create_weiterberechnung_document(self) -> None:
        """ Builds and creates the AQS Bericht. Therefore gathers all data from FE
        """

        global STATUS_MSG

        data = {
            "kennung_des_empfaengers": SELECTED_NACHWEIS["fakt_rechempf"].values[0],
            "ansprechpartner": SELECTED_NACHWEIS["ve_ansprechpartner"].values[0],
            "bestellnummer": SELECTED_NACHWEIS["fakt_bestellnummer"].values[0],
            "nachweisnummer": SELECTED_NACHWEIS["en_nachweisnummer"].values[0],
            "datum_probenahme": self.pnp_input_date_edit.text(),
            "material_name": SELECTED_NACHWEIS["ve_abfallbezeichnung"].values[0],
            "erzeuger_name": self.person_lineedit.text(),
            "projekt_nr": SELECTED_NACHWEIS["fakt_projektnummer"].values[0]
        }
        self.create_word(CONFIG_HELPER.get_specific_config_value("weiterberechnung_vorlage"), data, "Weiterberechnungsformular")

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
            file = QFileDialog.getSaveFileName(
                self, f'Speicherort für {dialog_file}', STANDARD_SAVE_PATH, filter='*.docx')
            if file[0]:
                self.word_helper.write_to_word_file(context=data, vorlage_path=vorlage, name=self.filename_from_path(file[0]))
                self.feedback_message(
                    "success", ["Das Protokoll wurde erfolgreich erstellt."])
                return file[0]
            else:
                print(file)
                return "/"
        except Exception as ex:
            STATUS_MSG = [
                f"{dialog_file} konnte nicht erstellt werden: " + str(ex)]
            self.feedback_message("error", STATUS_MSG)

    def filename_from_path(self, path):
        file_name = os.path.basename(path)
        return file_name

    def empty_all_nachweis_le(self) -> None:
        """ Empties all LineEdits in the Nachweis two Navigations
        """

        self.name_lineedit.setText("")
        self.person_lineedit.setText("")
        self.location_lineedit.setText("")
        self.avv_lineedit.setText("")
        self.amount_lineedit.setText("")


    def empty_values(self) -> None:
        """ Empties all LineEdits in the first two Navigations
        """

        self.empty_all_nachweis_le()

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

        ## PNP-In
        self.haufwerk_check.setChecked(False)
        self.bigbags_check.setChecked(False)
        self.silo_check.setChecked(False)

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

    def insert_values(self, id:str, origin: str= "nachweis") -> None:
        """ Inserts all value into CapZa FE based on selected Pobe
        """

        global STATUS_MSG

        # activate the analysebutton:
        self.nav_analysis_btn.setEnabled(True)

        self.empty_values()
        db_header_dict = DATABASE_HELPER.get_all_heading_names()
        STATUS_MSG = []
        # in Dateneingabe
        try:
            self.insert_nachweis_values()
        except Exception as ex:
            raise ex
        # in Analysewerte
        try:
            self.ph_lineedit.setText(str(round_if_psbl(
                format_specific_insert_value(db_header_dict["ph_wert"], SELECTED_PROBE))))
            try:
                if float(SELECTED_PROBE[db_header_dict["ph_wert"]]) <= 8:
                    self.set_ampel_color(self.ph_ampel_lbl, "yellow")
            except:
                pass
        except Exception as ex:
            STATUS_MSG.append(
                f"Der pH-Wert kann nicht interpretiert werden: [{ex}]")
        try:
            self.leitfaehigkeit_lineedit.setText(str(round_if_psbl(
                format_specific_insert_value(db_header_dict["leitfaehigkeit"], SELECTED_PROBE))))
            try:
                if float(SELECTED_PROBE[db_header_dict["leitfaehigkeit"],]) >= 12:
                    self.set_ampel_color(self.lf_ampel_lbl, "yellow")
            except:
                pass
        except Exception as ex:
            STATUS_MSG.append(
                f"Der Leitfähigkeitswert [mS/cm] kann nicht interpretiert werden: [{ex}]")
        try:
            self.feuchte_lineedit.setText(str(round_if_psbl(
                format_specific_insert_value(db_header_dict["wassergehalt"], SELECTED_PROBE))))
        except Exception as ex:
            STATUS_MSG.append(
                f"Der Wassergehalt [%] kann nicht interpretiert werden: [{ex}]")
        try:
            self.chrome_vi_lineedit.setText(str(round_if_psbl(format_specific_insert_value(
                db_header_dict["Cr 205.560 (Aqueous-Axial-iFR)"], SELECTED_PROBE))))
            try:
                if float(SELECTED_PROBE[db_header_dict["Cr 205.560 (Aqueous-Axial-iFR)"]]) <= 7:
                    self.set_ampel_color(self.chrom_ampel_lbl, "green")
                elif float(SELECTED_PROBE[db_header_dict["Cr 205.560 (Aqueous-Axial-iFR)"]]) > 7:
                    self.set_ampel_color(self.chrom_ampel_lbl, "red")
            except:
                pass
        except Exception as ex:
            STATUS_MSG.append(
                f"Der Chromwert kann nicht interpretiert werden: [{ex}]")
        try:
            self.lipos_ts_lineedit.setText(str(round_if_psbl(format_specific_insert_value(
                db_header_dict["result_lipos_ts"], SELECTED_PROBE))))
            try:
                if float(SELECTED_PROBE[db_header_dict["result_lipos_ts"]]) <= 4:
                    self.set_ampel_color(self.ts_ampel_lbl, "green")
                elif float(SELECTED_PROBE[db_header_dict["result_lipos_ts"]]) > 4:
                    self.set_ampel_color(self.ts_ampel_lbl, "red")
            except:
                pass
        except Exception as ex:
            STATUS_MSG.append(
                f"Der Lipos TS [%] kann nicht interpretiert werden: [{ex}]")
        try:
            self.lipos_os_lineedit.setText(format_specific_insert_value(
                db_header_dict["result_lipos_fs"], SELECTED_PROBE))
        except Exception as ex:
            STATUS_MSG.append(
                f"Der Lipos FS [%] kann nicht interpretiert werden: [{ex}]")
        try:
            self.gluehverlust_lineedit.setText(str(round_if_psbl(
                format_specific_insert_value(db_header_dict["result_gv"], SELECTED_PROBE))))
            try:
                if float(SELECTED_PROBE[db_header_dict["result_gv"]]) <= 10:
                    self.set_ampel_color(self.GV_ampel_lbl, "green")
                elif float(SELECTED_PROBE[db_header_dict["result_gv"]]) > 10:
                    self.set_ampel_color(self.GV_ampel_lbl, "red")
            except:
                pass
        except Exception as ex:
            STATUS_MSG.append(
                f"Der Glühverlust-Wert [%] kann nicht interpretiert werden: [{ex}]")
        try:
            self.doc_lineedit.setText(str(round_if_psbl(
                format_specific_insert_value(db_header_dict["doc"], SELECTED_PROBE))))
            try:
                if float(SELECTED_PROBE[db_header_dict["doc"]]) <= 100:
                    self.set_ampel_color(self.doc_ampel_lbl, "green")
                elif 100 < float(SELECTED_PROBE[db_header_dict["doc"]]) <= 199:
                    self.set_ampel_color(self.doc_ampel_lbl, "purple")
                elif float(SELECTED_PROBE[db_header_dict["doc"]]) > 199:
                    self.set_ampel_color(self.doc_ampel_lbl, "red")
            except:
                pass
        except Exception as ex:
            STATUS_MSG.append(
                f"Der DOC-Wert [mg/L] kann nicht interpretiert werden: [{ex}]")
        try:
            self.tds_lineedit.setText(str(round_if_psbl(format_specific_insert_value(
                db_header_dict["result_tds_ges"], SELECTED_PROBE))))
            try:
                if float(SELECTED_PROBE[db_header_dict["result_tds_ges"]]) <= 10000:
                    self.set_ampel_color(self.tds_ampel_lbl, "green")
                elif 10000 < float(SELECTED_PROBE[db_header_dict["result_tds_ges"]]) < 20000:
                    self.set_ampel_color(self.tds_ampel_lbl, "purple")
                elif float(SELECTED_PROBE[db_header_dict["result_tds_ges"]]) > 20000:
                    self.set_ampel_color(self.tds_ampel_lbl, "red")
            except:
                pass
        except Exception as ex:
            STATUS_MSG.append(
                f"Der Wert 'TDS Gesamt gelöste Stoffe (mg/L)' kann nicht interpretiert werden: [{ex}]")
        try:
            self.mo_lineedit.setText(str(round_if_psbl(
                format_specific_insert_value(db_header_dict["molybdaen"], SELECTED_PROBE))))
            try:
                if float(SELECTED_PROBE[db_header_dict["molybdaen"]]) <= 3:
                    self.set_ampel_color(self.mo_ampel_lbl, "green")
                elif float(SELECTED_PROBE[db_header_dict["molybdaen"]]) > 3:
                    self.set_ampel_color(self.mo_ampel_lbl, "red")
            except:
                pass
        except Exception as ex:
            STATUS_MSG.append(
                f"Der Molybdän-Wert kann nicht interpretiert werden: [{ex}]")
        try:
            self.se_lineedit.setText(str(round_if_psbl(format_specific_insert_value(
                db_header_dict["Se 196.090 (Aqueous-Axial-iFR)"], SELECTED_PROBE))))
            try:
                if float(SELECTED_PROBE[db_header_dict["Se 196.090 (Aqueous-Axial-iFR)"],]) <= 0.7:
                    self.set_ampel_color(self.se_ampel_lbl, "green")
                elif float(SELECTED_PROBE[db_header_dict["Se 196.090 (Aqueous-Axial-iFR)"],]) > 0.7:
                    self.set_ampel_color(self.se_ampel_lbl, "red")
            except:
                pass
        except Exception as ex:
            STATUS_MSG.append(
                f"Der Se-Wert kann nicht interpretiert werden: [{ex}]")
        try:
            self.sb_lineedit.setText(str(round_if_psbl(format_specific_insert_value(
                db_header_dict["Sb 206.833 (Aqueous-Axial-iFR)"], SELECTED_PROBE))))
            try:
                if float(SELECTED_PROBE[db_header_dict["Sb 206.833 (Aqueous-Axial-iFR)"],]) <= 0.5:
                    self.set_ampel_color(self.sb_ampel_lbl, "green")
                elif float(SELECTED_PROBE[db_header_dict["Sb 206.833 (Aqueous-Axial-iFR)"],]) > 0.7:
                    self.set_ampel_color(self.sb_ampel_lbl, "red")
            except:
                pass
        except Exception as ex:
            STATUS_MSG.append(
                f"Der Sb-Wert kann nicht interpretiert werden: [{ex}]")
        try:
            self.fluorid_lineedit.setText(str(round_if_psbl(
                format_specific_insert_value(db_header_dict["fluorid"], SELECTED_PROBE))))
            try:
                if float(SELECTED_PROBE[db_header_dict["fluorid"]]) <= 50:
                    self.set_ampel_color(self.fluorid_ampel_lbl, "green")
                elif float(SELECTED_PROBE[db_header_dict["fluorid"]]) > 50:
                    self.set_ampel_color(self.fluorid_ampel_lbl, "red")
            except:
                pass
        except Exception as ex:
            STATUS_MSG.append(
                f"Der Fluorid-Wert [mg/L] kann nicht interpretiert werden: [{ex}]")
        try:
            self.chlorid_lineedit.setText(str(round_if_psbl(
                format_specific_insert_value(db_header_dict["chlorid"], SELECTED_PROBE))))
            try:
                if float(SELECTED_PROBE[db_header_dict["chlorid"]]) <= 2500:
                    self.set_ampel_color(self.chlorid_ampel_lbl, "green")
                elif 2500 < float(SELECTED_PROBE[db_header_dict["chlorid"]]) <= 4000:
                    self.set_ampel_color(self.chlorid_ampel_lbl, "purple")
                elif float(SELECTED_PROBE[db_header_dict["chlorid"]]) > 4000:
                    self.set_ampel_color(self.chlorid_ampel_lbl, "red")
            except:
                pass
        except Exception as ex:
            STATUS_MSG.append(
                f"Der Chlorid-Wert [mg/L] kann nicht interpretiert werden: [{ex}]")
        try:
            self.toc_lineedit.setText(str(round_if_psbl(
                format_specific_insert_value(db_header_dict["toc"], SELECTED_PROBE))))
            try:
                if float(SELECTED_PROBE[db_header_dict["toc"]]) <= 6:
                    self.set_ampel_color(self.toc_ampel_lbl, "green")
                elif float(SELECTED_PROBE[db_header_dict["toc"]]) > 6:
                    self.set_ampel_color(self.toc_ampel_lbl, "red")
            except:
                pass
        except Exception as ex:
            STATUS_MSG.append(
                f"Der TOC-Wert [%] kann nicht interpretiert werden: [{ex}]")
        try:
            self.ec_lineedit.setText(str(round_if_psbl(
                format_specific_insert_value(db_header_dict["ec"], SELECTED_PROBE))))
        except Exception as ex:
            STATUS_MSG.append(
                f"Der EC-Wert [%] kann nicht interpretiert werden: [{ex}]")
        try:
            if SELECTED_PROBE[db_header_dict["Pb"]] != "<LOD":
                self.pb_lineedit.setText(str(float(
                    SELECTED_PROBE[db_header_dict["Pb"]]) * 10000)) if SELECTED_PROBE[db_header_dict["Pb"]] != None else "-"
            else:
                self.pb_lineedit.setText("-")
        except Exception as ex:
            STATUS_MSG.append(
                f"Der Pb kann nicht interpretiert werden: [{ex}]")

        date = str(SELECTED_PROBE["datum"]).split()[0]
        date = date.split("-")
        y = date[0]
        m = date[1]
        d = date[2]
        self.probe_date.setDate(QDate(int(y), int(m), int(d)))
        self.check_start_date.setDate(QDate(int(y), int(m), int(d)))
        self.pnp_output_probenahmedatum.setDate(QDate(int(y), int(m), int(d)))
        self.pnp_input_date_edit.setDate(QDate(int(y), int(m), int(d)))

        self.nachweisnr_lineedit.setText(
            str(SELECTED_PROBE[db_header_dict["material_kenn"]]))

        # in PNP Input
        if origin == "nachweis":
            selected_btbdata = RAMSES_HELPER.btb_depends_on_kennung(RAMSES_CONN, SELECTED_NACHWEIS["en_nachweisnummer"].values[0])
            
            lagerung = None
            if selected_btbdata.shape[0] and any(selected_btbdata):
                lagerung = selected_btbdata.iloc[0]['anlieferungsformtest']
                print(lagerung)
                if lagerung == "lose":
                    self.haufwerk_check.setChecked(True)
                if lagerung == "Silo":
                    self.silo_check.setChecked(True)
                if lagerung == "Big Bag":
                    self.bigbags_check.setChecked(True)      
            else:
                # get selected project nr
                pnr = projectnr.ProjectNrData()
                selected_btbdata = pnr.get_data()

        self.pnp_in_erzeuger_lineedit.setText(
            str(list(SELECTED_NACHWEIS["en_name"])[0]))
        self.pnp_in_abfallart_textedit.setPlainText(self.format_avv_space_after_every_second(str(list(
            SELECTED_NACHWEIS["en_asn_info_feld"])[0])) + ", " + str(list(SELECTED_NACHWEIS["ve_abfallbez_betr_intern"])[0]))


        if len(STATUS_MSG) > 0:
            self.feedback_message(
                "attention", ["Ein oder mehr Werte konnten nicht interpretiert werden."])
        else:
            self.feedback_message("success", ["Probe erfolgreich geladen."])
        self.show_second_info(
            "Gehe zu 'Analysewerte', um die Dokumente zu erstellen. >")

    def insert_nachweis_values(self, so: bool = False) -> None:
        """ Inserts all value into CapZa FE based on selected Pobe
        """

        global STATUS_MSG

        self.empty_all_nachweis_le()
        STATUS_MSG = []
        # in Dateneingabe
        try:
            self.kennung_combo.setCurrentText(
                format_specific_insert_value("nachweis_letters", SELECTED_NACHWEIS))
            self.pnr_combo.setCurrentText(
                format_specific_insert_value("project_yr", SELECTED_NACHWEIS))
            self.kennung_lineedit.setText(
                format_specific_insert_value("nachweis_nr", SELECTED_NACHWEIS))
            self.project_nr_lineedit.setText(
                format_specific_insert_value("project_nr", SELECTED_NACHWEIS))
            self.name_lineedit.setText(format_specific_insert_value(
                "ve_abfallbez_betr_intern", SELECTED_NACHWEIS))
            self.person_lineedit.setText(
                format_specific_insert_value("en_name", SELECTED_NACHWEIS))
            self.location_lineedit.setText(str(list(SELECTED_NACHWEIS["en_plz"])[
                                           0]) + " " + str(list(SELECTED_NACHWEIS["en_ort"])[0]))
            self.avv_lineedit.setText(self.format_avv_space_after_every_second(
                str(list(SELECTED_NACHWEIS["en_asn_info_feld"])[0])))
            self.amount_lineedit.setText(
                SELECTED_NACHWEIS['menge'].values[0].replace(",", ".") if SELECTED_NACHWEIS['menge'].values[0] != "nan" else "-"
                )

            # TODO: Setzte Nachweisnummer in PNP Input


            if so:
                # Deacrivate other Analysedaten:
                self.nav_analysis_btn.setEnabled(False)

            if len(STATUS_MSG) > 0:
                self.feedback_message(
                    "attention", ["Ein oder mehr Werte konnten nicht interpretiert werden."])
            else:
                self.feedback_message("success", ["Nachweis erfolgreich geladen."])

        except Exception as ex:
            print("FEHLER = " + str(ex))
        # in Analysewerte

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
        else:
            color = "#ffffff"

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

        global ALL_DATA_PROBE, STATUS_MSG
        try:
            STATUS_MSG = []
            if ALL_DATA_PROBE == 0:
                data = DATABASE_HELPER.get_all_probes()
                self.open_pop_up_probe(data)
                ALL_DATA_PROBE = data
            else:
                self.open_pop_up_probe(ALL_DATA_PROBE)
        except Exception as ex:
            STATUS_MSG.append(
                f"Es konnten keine Daten ermittelt werden: [{str(ex)}]")
            self.feedback_message("error", STATUS_MSG)

    def read_all_nachweis(self) -> None:
        """ Read all Nachweise (from Ramses) and save it globally
        """
        self.open_pop_up_nachweis(ALL_DATA_NACHWEIS)

    def format_avv_space_after_every_second(self, avv_raw: str) -> str:
        """ Formats en_asn_info_feld Number: After every second charackter adds a space

        Args:
            avv_raw (str): en_asn_info_feld Number
                e.g.: '00000000'

        Returns:
            str: Formatted en_asn_info_feld Number
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
            msg = "/"
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

        self.check_for_errors()
        self.status_msg_btm.show()
        QTimer.singleShot(4000, lambda: self.status_msg_btm.hide())

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
        self.model.setHorizontalHeaderLabels(["Datum", "Materialkennung", "Zusatzinfos", "Feuchte [%]", "Lipos [%]",
                                             "Glühverlust [%]", "TDS [mg/L]", "Chlorid [mg/L]", "pH Wert", "Molybdän [mg/L]", "DOC [%]", "TOC [%]", "EC [%]"])
        for row, text in enumerate(ALL_DATA_PROBE):
            date_item = QtGui.QStandardItem(
                str(text["datum"]) if text["datum"] != None else "")
            kennung_item = QtGui.QStandardItem(
                str(text["material_kenn"]) if text["material_kenn"] != None else "")
            material_item = QtGui.QStandardItem(
                str(text["material_bez"]) if text["material_bez"] != None else "")
            wasser_item = QtGui.QStandardItem(str(round_if_psbl(
                text["wassergehalt"])) if text["wassergehalt"] != None else "")
            liposts_item = QtGui.QStandardItem(str(round_if_psbl(
                text["result_lipos_ts"])) if text["result_lipos_ts"] != None else "")
            gv_item = QtGui.QStandardItem(
                str(round_if_psbl(text["result_gv"])) if text["result_gv"] != None else "")
            tds_gesamt_item = QtGui.QStandardItem(str(round_if_psbl(
                text["result_tds_ges"])) if text["result_tds_ges"] != None else "")
            chlorid_item = QtGui.QStandardItem(
                str(round_if_psbl(text["chlorid"])) if text["chlorid"] != None else "")
            ph_wert_item = QtGui.QStandardItem(
                str(round_if_psbl(text["ph_wert"])) if text["ph_wert"] != None else "")
            mo_item = QtGui.QStandardItem(
                str(round_if_psbl(text["molybdaen"])) if text["molybdaen"] != None else "")
            doc_item = QtGui.QStandardItem(
                str(round_if_psbl(text["doc"])) if text["doc"] != None else "")
            toc_item = QtGui.QStandardItem(
                str(round_if_psbl(text["toc"])) if text["toc"] != None else "")
            ec_item = QtGui.QStandardItem(
                str(round_if_psbl(text["ec"])) if text["ec"] != None else "")

            self.model.setItem(row, 0, date_item)
            self.model.setItem(row, 1, kennung_item)
            self.model.setItem(row, 2, material_item)
            self.model.setItem(row, 3, wasser_item)
            self.model.setItem(row, 4, liposts_item)
            self.model.setItem(row, 5, gv_item)
            self.model.setItem(row, 6, tds_gesamt_item)
            self.model.setItem(row, 7, chlorid_item)
            self.model.setItem(row, 8, ph_wert_item)
            self.model.setItem(row, 9, mo_item)
            self.model.setItem(row, 10, doc_item)
            self.model.setItem(row, 11, toc_item)
            self.model.setItem(row, 12, ec_item)

        # filter proxy model
        self.filter_proxy_model = QtCore.QSortFilterProxyModel()
        self.filter_proxy_model.setSourceModel(self.model)
        self.filter_proxy_model.setFilterKeyColumn(1)  # second column
        self.la_table_view.setModel(self.filter_proxy_model)
        self.la_search_found_lbl.setText(
            f"{self.filter_proxy_model.rowCount()} Ergebnisse gefunden")

        self.laborauswertung_lineedit.textChanged.connect(self.apply_filter)

    def apply_filter(self, text):
        self.filter_proxy_model.setFilterRegExp(text)
        self.la_search_found_lbl.setText(
            f"{self.filter_proxy_model.rowCount()} Ergebnisse gefunden")

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
            if ALL_DATA_PROBE == 0 or ALL_DATA_PROBE == None:
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

    def edit_laborauswertung(self) -> None:
        """ Loads the selected Laborauswertung row and open the frame to edit
        """
        row = (self.la_table_view.selectionModel().currentIndex())
        # kennung=index.sibling(index.row(),1).data()
        try:
            selected_datum = row.sibling(row.row(), 0).data()
            selected_kennung = row.sibling(row.row(), 1).data()
            selected_material = row.sibling(row.row(), 2).data()

            data = {
                "datum": selected_datum,
                "material_kenn": selected_kennung,
                "material_bez": selected_material
            }

            if data["datum"] and data["material_kenn"] and data["material_bez"]:
                la = Laborauswertung(self, "edit")
                la.insert_values(data)
                la.show()
            else:
                raise Exception(
                    "Ein dringlicher Wert wurde nicht eingetragen, der zur eindeutigen Identifikation essenziell ist.")
        except Exception as ex:
            STATUS_MSG = [
                f"Das Datum, die Kennung und das Material müssen zwingend angegeben sein: [{ex}]"]
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
                    if row["material_kenn"].split()[0] not in self._ene and isinstance(int(row["material_kenn"].split()[1]), int):
                        self._ene.append(row["material_kenn"].split()[0])
                except:
                    pass
            return self._ene
        else:
            ALL_DATA_PROBE = DATABASE_HELPER.get_all_probes()
            for row in ALL_DATA_PROBE:
                try:
                    if row["material_kenn"].split()[0] not in self._ene and isinstance(int(row["material_kenn"].split()[1]), int):
                        self._ene.append(row["material_kenn"].split()[0])
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


class PopUp(QtWidgets.QMainWindow):
    def __init__(self, parent=None, popup_type="Probe"):
        super(PopUp, self).__init__(parent)
        uic.loadUi(r'./views/select_probe.ui', self)

        self.setWindowTitle(f"CapZa - Zasada - { __version__ } - Wähle {popup_type}")
        self.popup_title_lbl.setText(f"Wähle {popup_type}:")
        self.load_data_btn.setText(f"Lade {popup_type}")

        self.cancel_btn.clicked.connect(self.close_window)
        init_shadow(self.load_data_btn)
        init_shadow(self.cancel_btn)

    def init_nachweis_data(self, dataset: pd.DataFrame) -> None:
        """ Inputs all the Probe data into the TableWidget

        Args:
            dataset (pd.DAtaframe): 
        """

        self.load_data_btn.clicked.connect(self.load_nachweis_data)
        try:
            self.model = QtGui.QStandardItemModel(dataset.shape[0], 2)
            self.model.setHorizontalHeaderLabels(
                ["Nachweisnummer", "Erzeuger"])
            for index,row in dataset.iterrows():
                i = index # For error case
                shorter_bezeichnung = QtGui.QStandardItem(row["en_nachweisnummer"][0:3]+" "+row["en_nachweisnummer"][-4:] if row["en_nachweisnummer"] else "-")
                erzeuger_item = QtGui.QStandardItem(row["en_name"] if row["en_name"] else "-")
                self.model.setItem(index, 0, shorter_bezeichnung)
                self.model.setItem(index, 1, erzeuger_item)
            # filter proxy model
            self.filter_proxy_model = QtCore.QSortFilterProxyModel()
            self.filter_proxy_model.setSourceModel(self.model)
            self.filter_proxy_model.setFilterKeyColumn(0)  # second column
            self.tableView.setModel(self.filter_proxy_model)
            self.results_lbl.setText(
                f"{self.filter_proxy_model.rowCount()} Ergebnisse gefunden")

            self.probe_filter_lineedit.textChanged.connect(self.apply_filter)
        except Exception as ex:
            print("FEHLER in:" ,i, row["en_nachweisnummer"], row["ve_abfallbez_betr_intern"],row["en_eingangsbest_beh"], str(ex))

    def init_probe_data(self, dataset: list[dict]) -> None:
        """ Inputs all the Probe data into the TableWidget

        Args:
            dataset (list[dict]): List of dictionaries with Probevalues
        """

        self.load_data_btn.clicked.connect(self.load_probe_data)

        self.model = QtGui.QStandardItemModel(len(dataset), 3)
        self.model.setHorizontalHeaderLabels(
            ["Datum", "Materialkennung", "Materialbezeichnung"])
        for row, text in enumerate(dataset):
            date_item = QtGui.QStandardItem(text["datum"])
            material_item = QtGui.QStandardItem(text["material_bez"])
            kennung_item = QtGui.QStandardItem(text["material_kenn"])
            self.model.setItem(row, 0, date_item)
            self.model.setItem(row, 2, material_item)
            self.model.setItem(row, 1, kennung_item)
        # filter proxy model
        self.filter_proxy_model = QtCore.QSortFilterProxyModel()
        self.filter_proxy_model.setSourceModel(self.model)
        self.filter_proxy_model.setFilterKeyColumn(1)  # second column
        self.tableView.setModel(self.filter_proxy_model)
        self.results_lbl.setText(
            f"{self.filter_proxy_model.rowCount()} Ergebnisse gefunden")

        self.probe_filter_lineedit.textChanged.connect(self.apply_filter)

    def apply_filter(self, text):
        self.filter_proxy_model.setFilterRegExp(text)
        self.results_lbl.setText(
            f"{self.filter_proxy_model.rowCount()} Ergebnisse gefunden")

    def load_nachweis_data(self) -> None:
        """ Gets the selected Probe and closes the Probe window
        """

        global SELECTED_PROBE, SELECTED_NACHWEIS
        try:
            index = (self.tableView.selectionModel().currentIndex())
            kennung = index.sibling(index.row(), 0).data()
            erzeuger = index.sibling(index.row(), 1).data()
            
            full_bezeichnung, kennung, numbers = get_full_project_ene_number(kennung)
            nachweis_data = ALL_DATA_NACHWEIS[ALL_DATA_NACHWEIS['en_nachweisnummer'] == full_bezeichnung]
            SELECTED_NACHWEIS = nachweis_data
            self.check_in_nachweis((full_bezeichnung, kennung, numbers), True)
            self.parent().insert_nachweis_values(True)
            self.close_window()
        except Exception as ex:
            STATUS_MSG.append("Fehler beim Laden des Nachweises: " +str(ex))
            self.parent().feedback_message("error", STATUS_MSG)

    def load_probe_data(self) -> None:
        """ Gets the selected Probe and closes the Probe window
        """

        global SELECTED_PROBE, SELECTED_NACHWEIS
        try:
            index = (self.tableView.selectionModel().currentIndex())
            kennung = index.sibling(index.row(), 1).data()
            material = index.sibling(index.row(), 2).data()
            date = index.sibling(index.row(), 0).data()
            SELECTED_PROBE = DATABASE_HELPER.get_specific_probe(
                id=kennung, material=material, date=date)
            kind = self.differentiate_probe(SELECTED_PROBE["material_kenn"])

            if kind == "ene":
                full_bezeichnung, kennung, numbers = get_full_project_ene_number(SELECTED_PROBE["material_kenn"])
                nachweis_data = ALL_DATA_NACHWEIS[ALL_DATA_NACHWEIS['en_nachweisnummer'] == full_bezeichnung]
                SELECTED_NACHWEIS = nachweis_data
                self.check_in_nachweis((full_bezeichnung, kennung, numbers))
                self.parent().insert_values(SELECTED_PROBE["material_kenn"])
            elif kind == "project":
                self.check_in_projekt_nummer(SELECTED_PROBE["material_kenn"])
                self.parent().insert_values(SELECTED_PROBE["material_kenn"], "project")
            
            self.close_window()
        except Exception as ex:
            STATUS_MSG.append("Fehler beim Laden der Probedaten: " +str(ex))
            self.parent().feedback_message("attention", STATUS_MSG)

    def differentiate_probe(self, wert: str) -> None:
        """ Decidees based on the wert where to look for information

        Args:
            wert (str): Number from the selected Probe
                e.g.: 'ENE123', '22-0000', ...

        Raises:
            Exception: Error when loading information to the probe
        """

        global SELECTED_NACHWEIS

        if re.match(r"[0-9]+-[0-9]+", wert):
            # es ist eine Projektnummer
            return "project"
        elif re.match(r"[a-zA-Z]+\s[0-9]+", wert):
            # es ist eine XXX 000 Kennung
            return "ene"
        elif re.match("[a-zA-Z]+\s['I']+", wert):
            raise Exception("DK Proben wurden nicht implementiert.")
        else:
            SELECTED_NACHWEIS = 0
            raise Exception("Andere Proben wurden nicht implementiert.")


    def check_in_nachweis(self, kennung_tpl: tuple, from_nachweis: bool = False) -> None:
        """ Loads the Nachweis data from Übersicht Nachweise

        Args:
            projektnummer (str): Nachweis Nr.
                e.g.: ('ENE', '2054', 'ENE5R3822054)
        """

        global SELECTED_NACHWEIS, SELECTED_PROBE
        nachweis_data = ALL_DATA_NACHWEIS[ALL_DATA_NACHWEIS['en_nachweisnummer'] == kennung_tpl[0]]
        if not from_nachweis:
            nachweis_data['nachweis_letters'] = kennung_tpl[1]
            nachweis_data['nachweis_nr'] = kennung_tpl[2]
            nachweis_data['project_yr'] = "-"
            nachweis_data['project_nr'] = "-"
        else:
            nachweis_data['nachweis_letters'] = kennung_tpl[1]
            nachweis_data['nachweis_nr'] = kennung_tpl[2]
            nachweis_data['project_yr'] = "-"
            nachweis_data['project_nr'] = "-"

        SELECTED_NACHWEIS = nachweis_data

        amount_valid_value = 0
        menge = 0
        mengen_list = [SELECTED_NACHWEIS["ve_anfall_kjahr1"].values[0], SELECTED_NACHWEIS["ve_anfall_kjahr2"].values[0], SELECTED_NACHWEIS["ve_anfall_kjahr3"].values[0], SELECTED_NACHWEIS["ve_anfall_kjahr4"].values[0], SELECTED_NACHWEIS["ve_anfall_kjahr5"].values[0]]
        # if any(mengen_list):
        for val in mengen_list:
            if val:
                amount_valid_value += 1
                menge += float(val)
        menge /= amount_valid_value
        SELECTED_NACHWEIS['menge'] = str(menge)

    def check_in_projekt_nummer(self, projektnummer: str) -> None:
        """ Loads the Nachweis data from ProjektNr.

        Args:
            projektnummer (str): Project Nr.
                e.g.: "22-0000"
        """

        global ALL_DATA_PROJECT_NR
        global SELECTED_NACHWEIS
        global STATUS_MSG

        for sheet in ALL_DATA_PROJECT_NR:
            try:
                if str(projektnummer) in ALL_DATA_PROJECT_NR[sheet]["Projekt-Nr."].values:
                    projekt_data = ALL_DATA_PROJECT_NR[sheet][ALL_DATA_PROJECT_NR[sheet]['Projekt-Nr.'] == str(projektnummer)]
            except Exception as ex:
                STATUS_MSG = [
                    f"Probe mit Sheet {sheet} konnte nicht geladen werden: [{ex}]"]

        projekt_data["en_ort"] = projekt_data["Ort"]
        projekt_data["en_plz"] = ""
        projekt_data["en_name"] = projekt_data['Erzeuger']
        projekt_data["en_asn_info_feld"] = projekt_data['AVV']
        projekt_data["ve_abfallbez_betr_intern"] = projekt_data['Material']

        projekt_data['nachweis_letters'] = "-"
        projekt_data['nachweis_nr'] = "-"
        projekt_data['project_yr'] = projektnummer.split("-")[0]
        projekt_data['project_nr'] = projektnummer.split("-")[1]
        SELECTED_NACHWEIS = projekt_data
        SELECTED_NACHWEIS["menge"] = str(SELECTED_NACHWEIS['Menge [t/a]'].values[0])

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
        self.setWindowTitle(
            f"CapZa - Zasada - {__version__} - Fehlerbeschreibung")
        error_long_msg = "Es wurden mehrere Fehler gefunden: \n\n"
        if len(STATUS_MSG) > 1:
            for error in STATUS_MSG:
                error_long_msg += f"- {str(error)}\n"
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
        self.parent().check_for_errors()


class Laborauswertung(QtWidgets.QDialog):
    def __init__(self, parent=None, la_type="add"):
        super(Laborauswertung, self).__init__(parent)
        uic.loadUi(r'./views/laborauswertung.ui', self)
        self.setWindowTitle(
            f"CapZa - Zasada - {__version__} - Laborauswertung")

        init_shadow(self.form_frame_1)
        init_shadow(self.form_frame_2)
        init_shadow(self.form_frame_3)

        self.save_la_btn.setEnabled(False)

        self.la_type = la_type
        self.icp_model = None
        self.rfa_model = None

        self.shown = True
        self.show_calculation_frame_btn.clicked.connect(
            self.toggle_calculation_data)
        self.show_calculation_frame_btn.setText("-")
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
                # Berechnung result_ts:
                if self.la_feuchte_input.text():
                    ts = 100-float(self.la_feuchte_input.text())
                    self.la_result_ts.setText(str(ts))
            except Exception as ex:
                raise Exception(f"Fehler bei Berechung des TS: [{ex}]")

            try:
                # Berechnung Lipos TS %:(Auswage-tara)/(einwaage_sox_frisch/100)
                if self.la_lipos_auswaage_input.text() and self.la_lipos_tara_input.text() and self.la_einwaage_sox_input.text():
                    if self.la_lipos_auswaage_input.text() != "-" and self.la_lipos_tara_input.text() != "-" and self.la_einwaage_sox_input.text() != "-":
                        lipos_ts = (float(self.la_lipos_auswaage_input.text(
                        )) - (float(self.la_lipos_tara_input.text())))/(float(self.la_einwaage_sox_input.text()) / 100)
                        self.la_result_lipos_ts.setText(str(lipos_ts))
            except Exception as ex:
                raise Exception(f"Fehler bei Berechung des Lipos TS: [{ex}]")
            try:
                # Berechnung GV %:
                if self.la_gv_auwaage_input.text() and self.la_gv_tara_input.text() and self.la_gv_einwaage_input.text():
                    if self.la_gv_auwaage_input.text() != "-" and self.la_gv_tara_input.text() != "-" and self.la_gv_einwaage_input.text() != "-":
                        gv = 100 - (float(self.la_gv_auwaage_input.text()) - float(
                            self.la_gv_tara_input.text())) / float(self.la_gv_einwaage_input.text()) * 100
                        self.la_result_gv.setText(str(gv))
            except Exception as ex:
                raise Exception(f"Fehler bei Berechung des GVs: [{ex}]")

            try:
                # Berechnung TDS:
                if self.la_tds_auswaage_input.text() and self.la_tds_tara_input.text() and self.la_tds_einwaage_input.text():
                    if self.la_tds_auswaage_input.text() != "-" and self.la_tds_tara_input.text() != "-" and self.la_tds_einwaage_input.text() != "-":
                        tds = (float(self.la_tds_auswaage_input.text()) - float(self.la_tds_tara_input.text())) * (
                            1000 / float(self.la_tds_einwaage_input.text()) * 1000)
                        self.la_result_tds_gesamt.setText(str(tds))
            except Exception as ex:
                raise Exception(
                    f"Fehler bei Berechung des TDS Gesamt gelöste Stoffe: [{ex}]")

            try:
                # Berechnung AOC:
                if self.la_toc_input.text() and self.la_ec_input.text():
                    if self.la_toc_input.text() != "-" and self.la_ec_input.text() != "-":
                        aoc = (float(self.la_toc_input.text()) - float(self.la_ec_input.text()))
                        self.la_result_aoc.setText(str(aoc))
            except Exception as ex:
                raise Exception(
                    f"Fehler bei Berechung des AOC: [{ex}]")

            self.save_la_btn.setEnabled(True)

        except Exception as ex:
            STATUS_MSG.append(ex)
            self.parent().feedback_message("error", STATUS_MSG)

    def insert_values(self, data: dict) -> None:
        """ Inserts all values when editing

        Args:
            data (dict): selected date, kennung and material
        """

        all_db_data = DATABASE_HELPER.get_specific_probe(
            id=data["material_kenn"], material=data["material_bez"], date=data["datum"])
        y, m, d = all_db_data["datum"].split()[0].split("-")
        self.la_date_edit.setDate(QDate(int(y), int(m), int(d)))
        try:
            self.la_material_input.setText(
                format_specific_insert_value("material_bez", all_db_data))

            self.la_kennung_input.setText(
                format_specific_insert_value("material_kenn", all_db_data))
            # ---------------------------------------------#
            self.la_feuchte_input.setText(
                format_specific_insert_value("wassergehalt", all_db_data))
            self.la_fluorid_input.setText(
                format_specific_insert_value("fluorid", all_db_data))

            self.la_ph_input.setText(
                format_specific_insert_value("ph_wert", all_db_data))
            self.la_lf_input.setText(
                format_specific_insert_value("leitfaehigkeit", all_db_data))
            self.la_cl_input.setText(
                format_specific_insert_value("chlorid", all_db_data))
            self.la_cr_input.setText(format_specific_insert_value(
                "Cr 205.560 (Aqueous-Axial-iFR)", all_db_data))
            self.la_doc_input.setText(
                format_specific_insert_value("doc", all_db_data))
            self.la_mo_input.setText(
                format_specific_insert_value("molybdaen", all_db_data))
            self.la_toc_input.setText(
                format_specific_insert_value("toc", all_db_data))
            self.la_ec_input.setText(
                format_specific_insert_value("ec", all_db_data))
            # ----------------------------------------------#
            # Berechnungs Daten
            self.la_einwaage_sox_input.setText(
                format_specific_insert_value("einwaage_sox_frisch", all_db_data))
            self.la_lipos_tara_input.setText(
                format_specific_insert_value("lipos_tara", all_db_data))
            self.la_lipos_auswaage_input.setText(
                format_specific_insert_value("lipos_auswaage", all_db_data))
            self.la_gv_tara_input.setText(
                format_specific_insert_value("gv_tara", all_db_data))
            self.la_gv_einwaage_input.setText(
                format_specific_insert_value("gv_einwaage", all_db_data))
            self.la_gv_auwaage_input.setText(
                format_specific_insert_value("gv_auswaage", all_db_data))
            self.la_tds_tara_input.setText(
                format_specific_insert_value("tds_tara", all_db_data))
            self.la_tds_einwaage_input.setText(
                format_specific_insert_value("tds_einwaage", all_db_data))
            self.la_tds_auswaage_input.setText(
                format_specific_insert_value("tds_auswaage", all_db_data))
            # -------------------------------------------------#
            # Ergebnis Daten
            self.la_result_ts.setText(str(
                100-float(all_db_data["wassergehalt"])) if all_db_data["wassergehalt"] else "-")
            self.la_result_lipos_ts.setText(
                format_specific_insert_value("result_lipos_ts", all_db_data))
            self.la_result_gv.setText(
                format_specific_insert_value("result_gv", all_db_data))
            self.la_result_tds_gesamt.setText(
                format_specific_insert_value("result_tds_ges", all_db_data))
            # -----------------------------------------------#
            # Bemerkungen
            try:
                bemerkung_lst = all_db_data[r"strukt_bemerkung"].split(";") if all_db_data[r"strukt_bemerkung"] else all_db_data[r"strukt_bemerkung"]
                self.listWidget.addItems(bemerkung_lst) if all_db_data[r"strukt_bemerkung"] != None else "-"
            except Exception as ex:
                print("Fehler bei der Bemerkung:", ex)

            # -----------------------------------------------#
            # RFA Daten
            rfa_headers = [
                # 'As 189.042',
                'Hg 194.227',
                'Se 196.090',
                'Mo 202.030',
                'Cr 205.560',
                'Sb 206.833',
                'Zn 213.856',
                'Pb 220.353',
                'Cd 228.802',
                'Ni 231.604',
                'Ba 233.527',
                'Fe 259.940',
                'Ca 318.128',
                'Cu 324.754',
                'Al 394.401',
                'Ar 404.442'
            ]
            try:
                if self.la_type != "add":
                    model = QStandardItemModel()
                    row = []
                    model.setHorizontalHeaderLabels(rfa_headers)
                    for header in rfa_headers:
                        item = QStandardItem(all_db_data[header + " (Aqueous-Axial-iFR)"])
                        item.setEditable(False)
                        row.append(item)
                    model.appendRow(row)
                    self.rfa_result_table_view.setModel(model)
            except Exception as ex:
                print("NESTED FEHLER: ", ex)

            # ICP Data
            icp_headers = ["Pb", "Pb Error", "Ni", "Ni Error", "Sb", "Sb Error", "Sn", "Sn Error", "Cd", "Cd Error", "Cr", "Cr Error", "Cu", "Cu Error", "Fe", "Fe Error", "Ag", "Ag Error", "Al", "Al Error", "As", "As Error", "Au", "Au Error", "Ba", "Ba Error", "Bal", "Bal Error", "Bi", "Bi Error", "Ca", "Ca Error", "Cl", "Cl Error", "Co", "Co Error",
                           "K", "K Error", "Mg", "Mg Error", "Mn", "Mn Error", "Mo", "Mo Error", "Nb", "Nb Error", "P", "P Error", "Pd", "Pd Error", "Rb", "Rb Error", "S", "S Error", "Se", "Se Error", "Si", "Si Error", "Sr", "Sr Error", "Ti", "Ti Error", "Tl", "Tl Error", "V", "V Error", "W", "W Error", "Zn", "Zn Error", "Zr", "Zr Error", "Br", "Br Error"]
            try:
                if self.la_type != "add":
                    model = QStandardItemModel()
                    row = []
                    model.setHorizontalHeaderLabels(icp_headers)
                    for header in icp_headers:
                        item = QStandardItem(all_db_data[header])
                        item.setEditable(False)
                        row.append(item)
                    model.appendRow(row)
                    self.icp_result_table_view.setModel(model)
            except Exception as ex:
                print("NESTED FEHLER: ", ex)
        except Exception as ex:
            print(f"FEHLER: {ex}")

    def la_collect_inserted_values(self) -> None:
        """ Adds the new Laborauswertung entry to the database
        """

        global STATUS_MSG, ALL_DATA_PROBE

        STATUS_MSG = []
        try:
            datum = self.la_date_edit.date().toString("yyyy-MM-dd 00:00:00")
            material = self.la_material_input.text()
            kennung = self.la_kennung_input.text()
            # ---------------------------------------------#
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
            # ----------------------------------------------#
            # Berechnungs Daten
            einwaage_sox = self.la_einwaage_sox_input.text()
            lipos_tara = self.la_lipos_tara_input.text()
            lipos_auswaage = self.la_lipos_auswaage_input.text()
            gv_tara = self.la_gv_tara_input.text()
            gv_einwaage = self.la_gv_einwaage_input.text()
            gv_auswaage = self.la_gv_auwaage_input.text()
            tds_tara = self.la_tds_tara_input.text()
            tds_einwaage = self.la_tds_einwaage_input.text()
            tds_auswaage = self.la_tds_auswaage_input.text()
            # -------------------------------------------------#
            # Ergebnis Daten
            result_ts = self.la_result_ts.text()
            result_lipos_ts = self.la_result_lipos_ts.text()
            result_gv = self.la_result_gv.text()
            result_tds_gesamt = self.la_result_tds_gesamt.text()
            self.save_data = {
                "datum": datum,
                "material_bez": material,
                "material_kenn": kennung,
                # -------------
                "wassergehalt": feuchte,
                "fluorid": fluorid,
                "ph_wert": ph,
                "leitfaehigkeit": leitf,
                "chlorid": chlorid,
                "Cr 205.560 (Aqueous-Axial-iFR)": cr,
                "doc": doc,
                "Mo 202.030 (Aqueous-Axial-iFR)": mo,
                "toc": toc,
                "ec": ec,
                # -------------
                "einwaage_sox_frisch": einwaage_sox,
                "lipos_tara": lipos_tara,
                "lipos_auswaage": lipos_auswaage,
                "gv_tara": gv_tara,
                "gv_einwaage": gv_einwaage,
                "gv_auswaage": gv_auswaage,
                "tds_tara": tds_tara,
                "tds_einwaage": tds_einwaage,
                "tds_auswaage": tds_auswaage,
                # --------------------------------#
                r"result_ts": result_ts,
                r"result_lipos_ts": result_lipos_ts,
                r"result_gv": result_gv,
                r"result_tds_ges": result_tds_gesamt,
            }

            # Bemerkung Daten
            bemerkung_hist = [self.listWidget.item(
                i).text() for i in range(self.listWidget.count())]
            if self.bemerkung_input.text():
                if len(bemerkung_hist) > 0:
                    bemerkung = ";".join(bemerkung for bemerkung in bemerkung_hist) + \
                        f";[{TODAY_FORMAT_STRING}]: {self.bemerkung_input.text()}"
                else:
                    bemerkung = f"[{TODAY_FORMAT_STRING}]: {self.bemerkung_input.text()}"
                self.save_data['strukt_bemerkung'] = bemerkung

        except Exception as ex:
            STATUS_MSG.append(f"Fehler beim Auslesen der Daten: [{ex}]")
            self.parent().check_for_errors()
            return

    def la_edit_save(self) -> None:
        lab_data = self.la_set_icp_and_rfa()
        self.la_collect_inserted_values()
        self.save_data |= lab_data
        if self.save_data["datum"] != "":
            if self.save_data["material_bez"] != "":
                if self.save_data["material_kenn"] != "":
                    DATABASE_HELPER.edit_laborauswertung(
                        self.save_data, self.save_data["material_kenn"], self.save_data["datum"])
                    self.close()
                else:
                    mark_error_line(self.la_kennung_input, "QTextEdit")
            else:
                mark_error_line(self.la_material_input, "QTextEdit")
        else:
            mark_error_line(self.la_date_edit, "QDateEdit")

    def la_add_save(self) -> None:
        lab_data = self.la_set_icp_and_rfa()
        self.la_collect_inserted_values()
        self.save_data |= lab_data
        if self.save_data["datum"] != "":
            if self.save_data["material_bez"] != "" or self.save_data["material_kenn"] != "":
                if self.save_data["material_bez"] != "":
                    if self.save_data["material_kenn"] != "":
                        DATABASE_HELPER.add_laborauswertung(self.save_data)
                        self.close()
                    else:
                        mark_error_line(self.la_kennung_input, "QTextEdit")
                else:
                    mark_error_line(self.la_material_input, "QTextEdit")
            else:
                mark_error_line(self.la_material_input, "QTextEdit")
                mark_error_line(self.la_kennung_input, "QTextEdit")
        else:
            mark_error_line(self.la_date_edit, "QDateEdit")

    def get_value_from_model(self, table_view, row: int, col: int):
        return table_view.model().data(table_view.model().index(row, col), role=Qt.DisplayRole)

    def la_set_icp_and_rfa(self):
        # if self.rfa_table_widget
        rfa_data = None
        icp_data = None
        if self.rfa_model:
            # get selected
            index = self.rfa_table_widget.currentIndex()
            row = index.row()
            name = self.rfa_table_widget.model().data(
                self.rfa_table_widget.model().index(row, 3), role=Qt.DisplayRole)
            rfa_data = {
                # "As 189.042 (Aqueous-Axial-iFR)": self.get_value_from_model(self.rfa_table_widget, row, 1),
                "Hg 194.227 (Aqueous-Axial-iFR)": self.get_value_from_model(self.rfa_table_widget, row, 2),
                "Se 196.090 (Aqueous-Axial-iFR)": self.get_value_from_model(self.rfa_table_widget, row, 3),
                "Mo 202.030 (Aqueous-Axial-iFR)": self.get_value_from_model(self.rfa_table_widget, row, 4),
                "Cr 205.560 (Aqueous-Axial-iFR)": self.get_value_from_model(self.rfa_table_widget, row, 5),
                "Sb 206.833 (Aqueous-Axial-iFR)": self.get_value_from_model(self.rfa_table_widget, row, 6),
                "Zn 213.856 (Aqueous-Axial-iFR)": self.get_value_from_model(self.rfa_table_widget, row, 7),
                "Pb 220.353 (Aqueous-Axial-iFR)": self.get_value_from_model(self.rfa_table_widget, row, 8),
                "Cd 228.802 (Aqueous-Axial-iFR)": self.get_value_from_model(self.rfa_table_widget, row, 9),
                "Ni 231.604 (Aqueous-Axial-iFR)": self.get_value_from_model(self.rfa_table_widget, row, 10),
                "Ba 233.527 (Aqueous-Axial-iFR)": self.get_value_from_model(self.rfa_table_widget, row, 11),
                "Fe 259.940 (Aqueous-Axial-iFR)": self.get_value_from_model(self.rfa_table_widget, row, 12),
                "Ca 318.128 (Aqueous-Axial-iFR)": self.get_value_from_model(self.rfa_table_widget, row, 13),
                "Cu 324.754 (Aqueous-Axial-iFR)": self.get_value_from_model(self.rfa_table_widget, row, 14),
                "Al 394.401 (Aqueous-Axial-iFR)": self.get_value_from_model(self.rfa_table_widget, row, 15),
                "Ar 404.442 (Aqueous-Axial-iFR)": self.get_value_from_model(self.rfa_table_widget, row, 16),
            }

        if self.icp_model:
            # get selected
            index = self.icp_table.currentIndex()
            row = index.row()
            icp_data = {
                "Pb": self.get_value_from_model(self.icp_table, row, 12),
                "Pb Error": self.get_value_from_model(self.icp_table, row, 13),
                "Ni": self.get_value_from_model(self.icp_table, row, 14),
                "Ni Error": self.get_value_from_model(self.icp_table, row, 14),
                "Sb": self.get_value_from_model(self.icp_table, row, 15),
                "Sb Error": self.get_value_from_model(self.icp_table, row, 16),
                "Sn": self.get_value_from_model(self.icp_table, row, 17),
                "Sn Error": self.get_value_from_model(self.icp_table, row, 18),
                "Cd": self.get_value_from_model(self.icp_table, row, 19),
                "Cd Error": self.get_value_from_model(self.icp_table, row, 20),
                "Cr": self.get_value_from_model(self.icp_table, row, 21),
                "Cr Error": self.get_value_from_model(self.icp_table, row, 22),
                "Cu": self.get_value_from_model(self.icp_table, row, 23),
                "Cu Error": self.get_value_from_model(self.icp_table, row, 24),
                "Fe": self.get_value_from_model(self.icp_table, row, 25),
                "Fe Error": self.get_value_from_model(self.icp_table, row, 26),
                "Ag": self.get_value_from_model(self.icp_table, row, 27),
                "Ag Error": self.get_value_from_model(self.icp_table, row, 28),
                "Al": self.get_value_from_model(self.icp_table, row, 29),
                "Al Error": self.get_value_from_model(self.icp_table, row, 30),
                "As": self.get_value_from_model(self.icp_table, row, 31),
                "As Error": self.get_value_from_model(self.icp_table, row, 32),
                "Au": self.get_value_from_model(self.icp_table, row, 33),
                "Au Error": self.get_value_from_model(self.icp_table, row, 34),
                "Ba": self.get_value_from_model(self.icp_table, row, 35),
                "Ba Error": self.get_value_from_model(self.icp_table, row, 36),
                "Bal": self.get_value_from_model(self.icp_table, row, 37),
                "Bal Error": self.get_value_from_model(self.icp_table, row, 38),
                "Bi": self.get_value_from_model(self.icp_table, row, 39),
                "Bi Error": self.get_value_from_model(self.icp_table, row, 40),
                "Ca": self.get_value_from_model(self.icp_table, row, 41),
                "Ca Error": self.get_value_from_model(self.icp_table, row, 42),
                "Cl": self.get_value_from_model(self.icp_table, row, 43),
                "Cl Error": self.get_value_from_model(self.icp_table, row, 44),
                "Co": self.get_value_from_model(self.icp_table, row, 45),
                "Co Error": self.get_value_from_model(self.icp_table, row, 46),
                "K": self.get_value_from_model(self.icp_table, row, 47),
                "K Error": self.get_value_from_model(self.icp_table, row, 48),
                "Mg": self.get_value_from_model(self.icp_table, row, 49),
                "Mg Error": self.get_value_from_model(self.icp_table, row, 50),
                "Mn": self.get_value_from_model(self.icp_table, row, 51),
                "Mn Error": self.get_value_from_model(self.icp_table, row, 52),
                "Mo": self.get_value_from_model(self.icp_table, row, 53),
                "Mo Error": self.get_value_from_model(self.icp_table, row, 54),
                "Nb": self.get_value_from_model(self.icp_table, row, 55),
                "Nb Error": self.get_value_from_model(self.icp_table, row, 56),
                "P": self.get_value_from_model(self.icp_table, row, 57),
                "P Error": self.get_value_from_model(self.icp_table, row, 58),
                "Pd": self.get_value_from_model(self.icp_table, row, 59),
                "Pd Error": self.get_value_from_model(self.icp_table, row, 60),
                "Rb": self.get_value_from_model(self.icp_table, row, 61),
                "Rb Error": self.get_value_from_model(self.icp_table, row, 62),
                "S": self.get_value_from_model(self.icp_table, row, 63),
                "S Error": self.get_value_from_model(self.icp_table, row, 64),
                "Se": self.get_value_from_model(self.icp_table, row, 65),
                "Se Error": self.get_value_from_model(self.icp_table, row, 66),
                "Si": self.get_value_from_model(self.icp_table, row, 67),
                "Si Error": self.get_value_from_model(self.icp_table, row, 68),
                "Sr": self.get_value_from_model(self.icp_table, row, 69),
                "Sr Error": self.get_value_from_model(self.icp_table, row, 70),
                "Ti": self.get_value_from_model(self.icp_table, row, 71),
                "Ti Error": self.get_value_from_model(self.icp_table, row, 72),
                "Tl": self.get_value_from_model(self.icp_table, row, 73),
                "Tl Error": self.get_value_from_model(self.icp_table, row, 74),
                "V": self.get_value_from_model(self.icp_table, row, 75),
                "V Error": self.get_value_from_model(self.icp_table, row, 76),
                "W": self.get_value_from_model(self.icp_table, row, 77),
                "W Error": self.get_value_from_model(self.icp_table, row, 78),
                "Zn": self.get_value_from_model(self.icp_table, row, 79),
                "Zn Error": self.get_value_from_model(self.icp_table, row, 80),
                "Zr": self.get_value_from_model(self.icp_table, row, 81),
                "Zr Error": self.get_value_from_model(self.icp_table, row, 82),
                "Br": self.get_value_from_model(self.icp_table, row, 83),
                "Br Error": self.get_value_from_model(self.icp_table, row, 84)
            }

        if rfa_data and icp_data:
            lab_data = rfa_data | icp_data
        elif rfa_data and not icp_data:
            lab_data = rfa_data
        elif not rfa_data and icp_data:
            lab_data = icp_data
        else:
            lab_data = {}

        return lab_data

    def close_window(self) -> None:
        self.close()

    def import_icp_scan(self) -> str:
        file = QFileDialog.getOpenFileName(
            self, "ICP-Scan", "C://", "Excel Files (*.xlsx *.xls)")
        try:
            icp_list = get_icp_data(file[0])

            # insert values to tablewidget
            headers = ["Reading No", "Time", "Type", "Duration", "Units", "Sequence", "Flags", "SAMPLE", "LOCATION", "INSPECTOR", "MISC", "NOTE", "Pb", "Pb Error", "Ni", "Ni Error", "Sb", "Sb Error", "Sn", "Sn Error", "Cd", "Cd Error", "Cr", "Cr Error", "Cu", "Cu Error", "Fe", "Fe Error", "Ag", "Ag Error", "Al", "Al Error", "As", "As Error", "Au", "Au Error", "Ba", "Ba Error", "Bal", "Bal Error", "Bi",
                       "Bi Error", "Ca", "Ca Error", "Cl", "Cl Error", "Co", "Co Error", "K", "K Error", "Mg", "Mg Error", "Mn", "Mn Error", "Mo", "Mo Error", "Nb", "Nb Error", "P", "P Error", "Pd", "Pd Error", "Rb", "Rb Error", "S", "S Error", "Se", "Se Error", "Si", "Si Error", "Sr", "Sr Error", "Ti", "Ti Error", "Tl", "Tl Error", "V", "V Error", "W", "W Error", "Zn", "Zn Error", "Zr", "Zr Error", "Br", "Br Error"]

            self.icp_model = DictionaryTableModel(icp_list, headers)
            self.icp_table.setModel(self.icp_model)
        except:
            print("Fehler bei ICP Laden")

    def import_rfa_scan(self) -> str:
        file = QFileDialog.getOpenFileName(
            self, "RFA-Scan", "C://", "CSV Files (*.csv)")
        try:
            rfa_list = get_rfa_data(file[0])

            # insert values to tablewidget
            headers = ["#",
                       "ExtCal.Average.1",
                       "ExtCal.Average.2",
                       "ExtCal.Average.3",
                       "ExtCal.Average.4",
                       "ExtCal.Average.5",
                       "ExtCal.Average.6",
                       "ExtCal.Average.7",
                       "ExtCal.Average.8",
                       "ExtCal.Average.9",
                       "ExtCal.Average.10",
                       "ExtCal.Average.11",
                       "ExtCal.Average.12",
                       "ExtCal.Average.13",
                       "ExtCal.Average.14",
                       "ExtCal.Average.15",
                       "ExtCal.Average.16"

                       ]

            self.rfa_model = DictionaryTableModel(rfa_list, headers)
            self.rfa_table_widget.setModel(self.rfa_model)
        except:
            print("Fehler bei RFA Laden")


def init_shadow(widget) -> None:
    """Sets shadow to the given widget

    Args:
        widget (QWidget): QFrame, QButton, ....
    """

    effect = QGraphicsDropShadowEffect()
    effect.setOffset(0, 1)
    effect.setBlurRadius(8)
    widget.setGraphicsEffect(effect)


def mark_error_line(widget: QtWidgets, widget_art: str) -> None:
    widget.setStyleSheet("""
        border: 2px solid red;
    """)

    QTimer.singleShot(3000, lambda: _set_default_style(widget, widget_art))


def _set_default_style(widget: QtWidgets, widget_art: str) -> None:
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
        """ % (widget_art, widget_art))

def get_full_project_ene_number(nummer: str) -> str:
        """ Gets the whole , VNE, ... Projectnr. from the shortform

        Args:
            nummer (str): Shortform of the ENE Nr.
                e.g.: ENE 1234

        Returns:
            str: Entire ENE Nr.
                e.g.: ENE382981234
        """

        letters, numbers = nummer.split()
        
        # find if 'is' substring is present
        result = ALL_DATA_NACHWEIS["en_nachweisnummer"].str.match(pat = f"{letters}.*?{numbers}", na=False)

        if any(result):
            return str(ALL_DATA_NACHWEIS["en_nachweisnummer"][result[result].index[0]]), letters, numbers
        else:
            raise Exception("Es konnte keine Nachweisnummer gefunden werden")

def write_multiple_fields(fields: list, text:str) -> None:
    for field in fields:
        field.setText(text)

class DictionaryTableModel(QtCore.QAbstractTableModel):
    def __init__(self, data, headers):
        super(DictionaryTableModel, self).__init__()
        self._data = data
        self._headers = headers

    def data(self, index, role):
        if role == Qt.DisplayRole:
            # Look up the key by header index.
            column = index.column()
            column_key = self._headers[column]
            return self._data[index.row()][column_key]

    def rowCount(self, index):
        # The length of the outer list.
        return len(self._data)

    def columnCount(self, index):
        # The length of our headers.
        return len(self._headers)

    def headerData(self, section, orientation, role):
        # section is the index of the column/row.
        if role == Qt.DisplayRole:
            if orientation == Qt.Horizontal:
                return str(self._headers[section])

            if orientation == Qt.Vertical:
                return str(section)


if __name__ == "__main__":
    STATUS_MSG = []
    lockfile_main = QtCore.QLockFile(QtCore.QDir.tempPath() + 'capza.lock')

    if lockfile_main.tryLock(100):
        app = QtWidgets.QApplication(sys.argv)
        # Create and display the splash screen
        splash_pix = QPixmap(":/logos/capza_logo.png").scaledToWidth(800)
        splash = QSplashScreen(splash_pix, Qt.WindowStaysOnTopHint)
        splash.setMask(splash_pix.mask())
        splash.show()

        d = {}
        try:
            d = CONFIG_HELPER.get_all_config()
        except Exception as ex:
            print(ex)
        if d:
            PNR_PATH = d["project_nr_path"]
            STANDARD_SAVE_PATH = d["save_path"]
            DB_PATH = d["db_path"]

        try:
            DATABASE_HELPER = DatabaseHelper(DB_PATH)
        except Exception as ex:
            STATUS_MSG.append(
                f"Die Datenbank konnte nicht gefunden werden. Bitte überprüfe in der Referenzeinstellung: [{ex}]")
        try:
            RAMSES_HELPER = RamsesHelper()
            RAMSES_CONN = RAMSES_HELPER.connect()
            ALL_DATA_NACHWEIS = RAMSES_HELPER.nachweis_data(RAMSES_CONN)
            if ALL_DATA_NACHWEIS is None:
                raise Exception
        except Exception as ex:
            STATUS_MSG.append(
                f"Es konnte keine Verbindung zu Ramses hergestellt werden. Überprüfe die Internetverbindung oder wende Dich an den Support")
        win = Ui()
        try:
            ALL_DATA_PROBE = DATABASE_HELPER.get_all_probes()
        except Exception as ex:
            print(ex)
            STATUS_MSG.append(
                f"Es konnten keine Proben geladen werden: [{ex}]")

        if STATUS_MSG:
            print(STATUS_MSG)
            win.feedback_message("error", STATUS_MSG)

        win.check_for_errors()

        splash.finish(win)
        win.show()
        win.check_for_update()
        sys.exit(app.exec_())
    else:
        print("Datei bereits geöffnet")
