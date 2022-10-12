
from pickle import GLOBAL
import pandas as pd
from PyQt5 import QtWidgets, uic, QtGui
from threading import Thread
from PyQt5.QtWidgets import QFileDialog, QTableWidgetItem
import sys

SELECTED_PROBE = 0

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
        self.excel_data = pd.read_excel(self.file)

    def insert_values(self):
        global SELECTED_PROBE

        self.project_nr_lineedit.setText(SELECTED_PROBE["Kennung \nDiese Zeile wird zum Sortieren benötigt"])


    def open_specific_probe(self, checked):
        # win_thread = Thread(target=self.open_probe_win)
        # read_excel_thread = Thread(target=self.read_excel)
        # win_thread.start()
        # read_excel_thread.start()
        self.read_excel()
        self.open_probe_win(self.excel_data)
        

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