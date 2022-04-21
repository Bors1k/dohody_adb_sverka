from PyQt5 import QtWidgets, QtCore
from ui_windows.MainWindow import Ui_MainWindow

class MainForm(QtWidgets.QMainWindow):
    def __init__(self):
        super(MainForm, self).__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.ui.pbutton_choose_files.clicked.connect(self.choose_starter_file)


    def choose_starter_file(self):
        self.qfiledlg = QtWidgets.QFileDialog()
        self.qfiledlg.setFileMode(QtWidgets.QFileDialog.FileMode.ExistingFile)
        file_path, _ = self.qfiledlg.getOpenFileName(self,filter='Excel (*.xls *.xlsx)')
        splitted = file_path.split("/")
        len = len(splitted)
        # file_name = 
        
        # if self.qfiledlg.exec() == QtWidgets.QDialog.DialogCode.Accepted:
        #     path = str(self.qfiledlg.selectedFiles()[0])
        #     print(path)