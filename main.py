
from pickle import GLOBAL
import pandas as pd
from PyQt5 import QtWidgets, uic, QtGui
from threading import Thread
from PyQt5.QtWidgets import QFileDialog, QTableWidgetItem
import sys

SELECTED_PROBE = 0
ALL_DATA = 0

class Ui(QtWidgets.QMainWindow):
    def __init__(self, parent=None):
        super(Ui, self).__init__(parent) # Call the inherited classes __init__ method
        uic.loadUi('./views/main.ui', self) # Load the .ui file
        self.show() # Show the GUI



        self.setWindowTitle("CapZa - v0.1 - Zasada") 


        self.stackedWidget.setCurrentIndex(0)
        self.status_msg_label.setText("")
        self.file = ""

        self.nav_data_btn.clicked.connect(lambda : self.stackedWidget.setCurrentIndex(0))
        self.nav_analysis_btn.clicked.connect(lambda : self.stackedWidget.setCurrentIndex(1))
        self.nav_pnp_entry_btn.clicked.connect(lambda : self.stackedWidget.setCurrentIndex(2))
        self.nav_pnp_output_btn.clicked.connect(lambda : self.stackedWidget.setCurrentIndex(3))
        self.nav_order_form_btn.clicked.connect(lambda : self.stackedWidget.setCurrentIndex(4))

        self.find_excel_btn.clicked.connect(self.select_excel)
        self.select_probe_btn.clicked.connect(self.open_specific_probe)


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


    def read_excel(self):
        excel_raw = pd.read_excel(self.file)
        return excel_raw

    def insert_values(self):
        global SELECTED_PROBE
        print(SELECTED_PROBE)

        ### in Dateneingabe
        self.project_nr_lineedit.setText(str(SELECTED_PROBE["Kennung \nDiese Zeile wird zum Sortieren benötigt"]))
        self.name_lineedit.setText(SELECTED_PROBE["Material beliebige Bezeichnung"])
        self.person_lineedit.setText(SELECTED_PROBE["Material beliebige Bezeichnung"])
        self.location_lineedit.setText(SELECTED_PROBE["Material beliebige Bezeichnung"])
        self.avv_lineedit.setText(SELECTED_PROBE["Material beliebige Bezeichnung"])
        self.amount_lineedit.setText(SELECTED_PROBE["Material beliebige Bezeichnung"])
        self.probe_amount_lineedit.setText(SELECTED_PROBE["Material beliebige Bezeichnung"])
        self.color_lineedit.setText(SELECTED_PROBE["Material beliebige Bezeichnung"])
        self.consistency_lineedit.setText(SELECTED_PROBE["Material beliebige Bezeichnung"])
        self.smell_lineedit.setText(SELECTED_PROBE["Material beliebige Bezeichnung"])
        self.remark_textedit.setText(SELECTED_PROBE["Material beliebige Bezeichnung"])

        ### in Analysewerte
        self.ph_lineedit.setText(str(SELECTED_PROBE["pH-Wert"]))
        self.leitfaehigkeit_lineedit.setText(str(SELECTED_PROBE["Leitfähigkeit (mS/cm)"]))
        self.feuchte_lineedit.setText(str(SELECTED_PROBE["Wassergehalt %"]))
        self.chrome_vi_lineedit.setText(str(SELECTED_PROBE["Cr 205.560 (Aqueous-Axial-iFR)"]))
        self.lipos_ts_lineedit.setText(str(SELECTED_PROBE["Lipos TS\n%"]))
        self.lipos_os_lineedit.setText(str(SELECTED_PROBE["Lipos FS\n%"]))
        self.gluehverlus_lineedit.setText(str(SELECTED_PROBE["GV [%]"]))
        self.doc_lineedit.setText(str(SELECTED_PROBE["Bezogen auf das eingewogene Material DOC mg/L "]))
        self.tds_lineedit.setText(str(SELECTED_PROBE["\nTDS\nGesamt gelöste Stoffe (mg/L)"]))
        self.mo_lineedit.setText(str(SELECTED_PROBE[" Bezogen auf das eingewogene Material Molybdän mg/L ………"]))
        self.se_lineedit.setText(str(SELECTED_PROBE["Se 196.090 (Aqueous-Axial-iFR)"]))
        self.sb_lineedit.setText(str(SELECTED_PROBE["Sb 206.833 (Aqueous-Axial-iFR)"]))
        self.fluorid_lineedit.setText(str(SELECTED_PROBE["Fluorid mg/L"]))
        self.chlorid_lineedit.setText(str(SELECTED_PROBE["Chlorid mg/L"]))
        # self.nh3_lineedit.setText("")
        # self.h2_lineedit.setText()
        self.toc_lineedit.setText(str(SELECTED_PROBE["TOC\n%"]))
        self.ec_lineedit.setText(str(SELECTED_PROBE["EC\n%"]))


    def open_specific_probe(self, checked):
        global ALL_DATA

        print(type(ALL_DATA))

        if isinstance(ALL_DATA, int):
            print("INTEGER")
            data = self.read_excel()
            print(type(data))
            self.open_probe_win(data)
            ALL_DATA = data
        else:
            print("kein Integer")
            self.open_probe_win(ALL_DATA)

        

        # win_thread.join()
        # read_excel_thread.join()
        

    def set_status(self, msg):
        self.status_msg_label.setText(msg)





class Probe(QtWidgets.QMainWindow): 
    def __init__(self, parent=None):
        super(Probe, self).__init__(parent)
        uic.loadUi('./views/select_probe.ui', self)
        
        
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
        self.parent().insert_values()
        self.close_window()

    def close_window(self):
        self.hide()



if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    win = Ui()
    sys.exit(app.exec_())