from PyQt5 import QtCore, QtGui, QtWidgets
class Worker(QtCore.QThread):

    progressMade = QtCore.pyqtSignal(int)
    finished = QtCore.pyqtSignal(int)
    interrupted = QtCore.pyqtSignal()

    def __init__(self, movie_groups):
        super().__init__()
        self.movie_groups = movie_groups

    def run(self):
        for i, (basename, paths) in enumerate(self.movie_groups.items()):
            io.to_raw_combined(basename, paths)
            self.progressMade.emit(i + 1)
        self.finished.emit(i)