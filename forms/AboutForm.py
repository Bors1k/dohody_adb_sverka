from PyQt5.QtWidgets import QMessageBox
from PyQt5 import QtWidgets

from ui_windows.AboutWindow import Ui_aboutWindow
from constants import app_constants



class AboutForm(QtWidgets.QWidget):
    """Класс окна о программе"""
    def __init__(self):
        super(AboutForm, self).__init__()
        self.ui = Ui_aboutWindow()
        self.ui.setupUi(self)
        self.ui.poLabel.setText(app_constants.app_name)
        self.ui.versionLabel.setText(self.ui.versionLabel.text() + "v" + app_constants.app_version)