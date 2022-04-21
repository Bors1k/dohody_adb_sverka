from PyQt5 import QtWidgets, QtCore
from libs.read_excel import ReadExcel
from ui_windows.MainWindow import Ui_MainWindow
from forms.AboutForm import AboutForm
import re, os
from constants import app_constants

class MainForm(QtWidgets.QMainWindow):
    def __init__(self):
        super(MainForm, self).__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.ui.action.triggered.connect(self.open_about_form)
        self.setWindowTitle(app_constants.app_name + " v" + app_constants.app_version)
        self.ui.pbutton_choose_files.clicked.connect(self.choose_starter_file)
        self.ui.txt_edit_protokol.setReadOnly(True)
        self.files_path = None
        self.files = None
        self.about_form = None

        self.read_excel_thread = None
    
    def open_about_form(self):
        self.about_form = AboutForm()
        self.about_form.show()

    @QtCore.pyqtSlot(str)
    def add_message_to_protokol(self, message):
        self.ui.txt_edit_protokol.append(message)

    def choose_starter_file(self):
        self.ui.txt_edit_protokol.clear()
        self.qfiledlg = QtWidgets.QFileDialog()
        self.qfiledlg.setFileMode(QtWidgets.QFileDialog.FileMode.ExistingFile)
        file_path, _ = self.qfiledlg.getOpenFileName(self,filter='Excel (*.xls *.xlsx)')
        if file_path != '':
            splitted = file_path.split("/")
            length = len(splitted)
            file_name = splitted[length-1]
            self.files_path = file_path[:-len(file_name)]
            ubp = file_name[1:6]

            regular = re.compile(r"^C?V?O?"+ubp+r"\.xls.?$")

            self.files = {}

            for file_name in os.listdir(self.files_path):
                if(os.path.isfile(self.files_path + file_name)):
                    if regular.match(file_name) is not None:
                        self.files[file_name[0]] = file_name

            if len(self.files) < 3:
                print("Не хватает файлов")
            elif len(self.files) > 3:
                print("В папке присутствует слишком много файлов с данным УБП")
            else:
                if self.check_current_files() == True:
                    self.read_excel_thread = ReadExcel(self, self.files, self.files_path, ubp)
                    self.read_excel_thread.protokol.connect(self.add_message_to_protokol)
                    self.read_excel_thread.start()
                    print("Все файлы получены")

                else:
                    print("Нет всех нужных файлов")
                    for file in self.files:
                        print(file)


    def check_current_files(self):
        c_count = 0
        v_count = 0
        o_count = 0

        for type, file in self.files.items():

            if type.__contains__("C"):
                c_count+=1
            if type.__contains__("V"):
                v_count+=1
            if type.__contains__("O"):
                o_count+=1

        if c_count == 1 and v_count == 1 and o_count == 1:
            return True
        return False