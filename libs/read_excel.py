import os
from PyQt5.QtCore import QThread, pyqtSignal, QSettings
import pandas as pd

class ReadExcel(QThread):
    def __init__(self, my_window, mass_excel, file_path):
       super(ReadExcel, self).__init__()
       self.my_window = my_window
       self.work = False
       self.mass_excel = mass_excel
       self.file_path = file_path

    def run(self):
        # Выписка из лицевого счета
        v_workbook = pd.read_html(self.file_path + self.mass_excel['V'])
        frame = v_workbook[0]
        values = frame.values

        v_postup_col = 0
        V_vozvrat_col = 0
        v_zachet_col = 0

        v_postup_value = ""
        v_vozvrat_value = ""
        v_zachet_value = ""

        for i in range(len(values)):
            for j in range(len(values[i])):
                if str(values[i][j]).__contains__("Поступления"):
                    v_postup_col = j
                if str(values[i][j]).__contains__("Возвраты"):
                    v_vozvrat_col = j
                if str(values[i][j]).__contains__("Зачеты"):
                    v_zachet_col = j
                if str(values[i][j]).__contains__("на конец дня"):
                    v_postup_value = values[i][v_postup_col]
                    v_vozvrat_value = values[i][v_vozvrat_col]
                    v_zachet_value = values[i][v_zachet_col]


        print(v_postup_value)
        print(v_vozvrat_value)
        print(v_zachet_value)

        # Отчет о состоянии лицевого счета
        
        