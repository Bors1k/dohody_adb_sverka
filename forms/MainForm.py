from PyQt5 import QtWidgets, QtCore
from PyQt5.QtWidgets import QFileDialog
from libs.read_excel import ReadExcel
from libs.save_to_excel import SaveExcel
from ui_windows.MainWindow import Ui_MainWindow
from forms.AboutForm import AboutForm
import re, os
import openpyxl, xlrd, lxml
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
        self.ui.pb_save_to_excel.setVisible(False)
        self.ui.pb_save_to_excel.clicked.connect(self.save_to_excel)

        self.files_path = None
        self.files = None
        self.about_form = None

        self.read_excel_thread = None
        self.messages = []
    
    def open_about_form(self):
        self.about_form = AboutForm()
        self.about_form.show()

    @QtCore.pyqtSlot(str)
    def add_message_to_protokol(self, message):
        self.ui.txt_edit_protokol.append(message)

    def choose_starter_file(self):
        self.ui.txt_edit_protokol.clear()
        self.messages = []
        self.qfiledlg = QtWidgets.QFileDialog()
        self.qfiledlg.setFileMode(QtWidgets.QFileDialog.FileMode.ExistingFile)
        file_path, _ = self.qfiledlg.getOpenFileName(self,filter='Excel (*.xls *.xlsx)')
        if file_path != '':
            splitted = file_path.split("/")
            length = len(splitted)
            file_name = splitted[length-1]
            self.files_path = file_path[:-len(file_name)]

            ubps = set()
            regular_for_files_ubp = re.compile(r"^[CVO](.{5,8})\.xls.?$")
            
            for file_name in os.listdir(self.files_path):
                if(os.path.isfile(self.files_path + file_name)):
                    if regular_for_files_ubp.match(file_name) is not None:
                        ubps.add(regular_for_files_ubp.match(file_name).groups(0)[0])
                
           

            self.files = {}
            for file_name in os.listdir(self.files_path):
                if(os.path.isfile(self.files_path + file_name)):
                    for ubp in ubps:
                        regular = re.compile(r"^[CVO]"+ubp+"\.xls.?$")
                        if regular.match(file_name) is not None:
                            try:
                                self.files[ubp][file_name[0]] = file_name
                            except KeyError as ex:
                                self.files[ubp] = {}
                                self.files[ubp][file_name[0]] = file_name


            print(self.files)
            temp_files = self.files.copy()
            for ubp, files in self.files.items():
                flag, message = self.check_current_files(files, ubp)
                if flag == False:
                    self.add_message_to_protokol("<p>Для УБП {} <font color='red'>{}</font></p>".format(ubp,message))
                    del temp_files[ubp]

            self.files = temp_files
            print(self.files)

            if len(self.files) != 0:
                self.read_excel_thread = ReadExcel(self, self.files, self.files_path)
                self.read_excel_thread.protokol.connect(self.add_message_to_protokol)
                self.read_excel_thread.pb_change_visible.connect(self.change_visible_excel_button)
                self.read_excel_thread.excel_protokol.connect(self.append_messages)
                self.read_excel_thread.start()
            else:
                self.add_message_to_protokol("<p>Не найдено файлов для проверки!</p>")

    @QtCore.pyqtSlot(dict)
    def append_messages(self, dict):
        self.messages.append(dict)

    @QtCore.pyqtSlot()
    def change_visible_excel_button(self):
        self.ui.pb_save_to_excel.setVisible(not self.ui.pb_save_to_excel.isVisible())

    def save_to_excel(self):
        filename, _ = QFileDialog.getSaveFileName(self,'Сохранить файл','Протокол.xlsx', ".xlsx(*.xlsx)")
        print(filename)
        self.save_to_excel_thread = SaveExcel(self.messages, filename)
        self.save_to_excel_thread.pb_change_visible.connect(self.change_visible_excel_button)
        self.save_to_excel_thread.start()

    def check_current_files(self, dict, ubp):
        c_count = 0
        v_count = 0
        o_count = 0

        for type, file in dict.items():

            if type.__contains__("C"):
                c_count+=1
            if type.__contains__("V"):
                v_count+=1
            if type.__contains__("O"):
                o_count+=1

        if c_count == 1 and v_count == 1 and o_count == 1:
            return True, ""

        message = ""
        if c_count == 0:
            message += "Не могу найти файл Справки C{}.xls \n".format(ubp)
        if v_count == 0:
            message += "Не могу найти файл Отчета O{}.xls \n".format(ubp)
        if v_count == 0:
            message += "Не могу найти файл Выписки V{}.xls \n".format(ubp)

        return False, message