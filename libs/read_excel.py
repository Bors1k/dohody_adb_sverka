import os
from PyQt5.QtCore import QThread, pyqtSignal, QSettings
from PyQt5 import QtCore
import pandas as pd

class ReadExcel(QThread):
    protokol = QtCore.pyqtSignal(str)


    def __init__(self, my_window, mass_excel, file_path, ubp):
       super(ReadExcel, self).__init__()
       self.my_window = my_window
       self.work = False
       self.mass_excel = mass_excel
       self.file_path = file_path
       self.ubp = ubp

    def run(self):
        # Выписка из лицевого счета АДБ файл «VXXXXX. xls»
        v_workbook = pd.read_html(self.file_path + self.mass_excel['V'])
        frame = v_workbook[0]
        values = frame.values

        v_postup_col = 0
        v_vozvrat_col = 0
        v_zachet_col = 0

        v_postup_value = 0
        v_vozvrat_value = 0
        v_zachet_value = 0

        for i in range(len(values)):
            for j in range(len(values[i])):
                if str(values[i][j]).__contains__("Поступления"):
                    v_postup_col = j
                if str(values[i][j]).__contains__("Возвраты"):
                    v_vozvrat_col = j
                if str(values[i][j]).__contains__("Зачеты"):
                    v_zachet_col = j
                if str(values[i][j]).__contains__("на конец дня"):
                    v_postup_value = float(values[i][v_postup_col].replace(' ','').replace(',','.'))
                    v_vozvrat_value = float(values[i][v_vozvrat_col].replace(' ','').replace(',','.'))
                    v_zachet_value = float(values[i][v_zachet_col].replace(' ','').replace(',','.'))


        print(str(v_postup_value))
        print(str(v_vozvrat_value))
        print(str(v_zachet_value))

        # Отчет о состоянии лицевого счета АДБ файл «OXXXXX. xls»
        o_workbook = pd.read_excel(self.file_path + self.mass_excel['O'], sheet_name=1)
        values = o_workbook.values

        o_postup_value = 0
        o_vozvrat_value = 0
        o_zachet_value = 0
        o_itogo_value = 0

        for i in range(len(values)):
            if str(values[i][1]) == "Итого:":
                o_postup_value = float(str(values[i][2]).replace(' ','').replace(',','.'))
                o_vozvrat_value = float(str(values[i][3]).replace(' ','').replace(',','.'))
                o_zachet_value = float(str(values[i][4]).replace(' ','').replace(',','.'))
                o_itogo_value = float(str(values[i][5]).replace(' ','').replace(',','.'))

        print(str(o_postup_value))
        print(str(o_vozvrat_value))
        print(str(o_zachet_value))
        print(str(o_itogo_value))

        # Справка о перечислении поступлений в бюджеты для АДБ файл «CXXXXX.xls»
        c_workbook = pd.read_excel("C:\\Users\\ufk4800_mnovikov\\Desktop\\TZ\\C00001.xls", sheet_name=1)
        values = c_workbook.values

        c_perechislen_value = 0
        c_ostatok_value = 0
        c_postup_vsego_value = 0

        for i in range(len(values)):
            if str(values[i][2]) == "Всего по разделам I и II":
                c_perechislen_value = float(str(values[i][7]).replace(' ','').replace(',','.'))


        c_workbook = pd.read_excel("C:\\Users\\ufk4800_mnovikov\\Desktop\\TZ\\C00001.xls", sheet_name=2)
        values = c_workbook.values

        for i in range(len(values)):
            if str(values[i][2]) == "Всего по разделам I и II":
                c_ostatok_value = float(str(values[i][14]).replace(' ','').replace(',','.'))

        c_workbook = pd.read_excel("C:\\Users\\ufk4800_mnovikov\\Desktop\\TZ\\C00001.xls", sheet_name=3)
        values = c_workbook.values

        for i in range(len(values)):
            if str(values[i][2]) == "Всего по разделу III":
                c_postup_vsego_value = float(str(values[i][3]).replace(' ','').replace(',','.'))

        print(str(c_perechislen_value))
        print(str(c_ostatok_value))
        print(str(c_postup_vsego_value))


        self.protokol.emit("Результат сверки для УБП {}".format(self.ubp))
        if v_postup_value == o_postup_value:
            self.protokol.emit("&nbsp;&nbsp;&nbsp;&nbsp;Поступления в выписке и отчете <font color='green'>Равны</font>")
        else:
            self.protokol.emit("&nbsp;&nbsp;&nbsp;&nbsp;Поступления в выписке и отчете <font color='red'>Различаются на {}</font>".format(abs(v_postup_value-o_postup_value)))
        
        if v_vozvrat_value == o_vozvrat_value:
            self.protokol.emit("&nbsp;&nbsp;&nbsp;&nbsp;Возвраты в выписке и отчете <font color='green'>Равны</font>")
        else:
            self.protokol.emit("&nbsp;&nbsp;&nbsp;&nbsp;Возвраты в выписке и отчете <font color='red'>Различаются на {}</font>".format(abs(v_vozvrat_value-o_vozvrat_value)))
        
        if v_zachet_value == o_zachet_value:
            self.protokol.emit("&nbsp;&nbsp;&nbsp;&nbsp;Зачеты в выписке и отчете <font color='green'>Равны</font>")
        else:
            self.protokol.emit("&nbsp;&nbsp;&nbsp;&nbsp;Зачеты в выписке и отчете <font color='red'>Различаются на {}</font>".format(abs(v_zachet_value-o_zachet_value)))
        
        if o_itogo_value == c_perechislen_value + c_ostatok_value + c_postup_vsego_value:
            self.protokol.emit("&nbsp;&nbsp;&nbsp;&nbsp;Итого в выписке и отчете <font color='green'>Равны</font>")
        else:
            self.protokol.emit("&nbsp;&nbsp;&nbsp;&nbsp;Поступления в выписке и отчете <font color='red'>Различаются на {}</font>".format(abs((c_perechislen_value + c_ostatok_value + c_postup_vsego_value)-o_itogo_value)))
        