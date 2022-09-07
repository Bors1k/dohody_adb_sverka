import os
from PyQt5.QtCore import QThread, pyqtSignal, QSettings
from PyQt5 import QtCore
from numpy import str_
import pandas as pd

class ReadExcel(QThread):
    protokol = QtCore.pyqtSignal(str)
    excel_protokol = QtCore.pyqtSignal(dict)
    pb_change_visible = QtCore.pyqtSignal()


    def __init__(self, my_window, mass_excel, file_path):
       super(ReadExcel, self).__init__()
       self.my_window = my_window
       self.work = False
       self.mass_excel = mass_excel
       self.file_path = file_path
       self.data = ''

    def run(self):
        for ubp, files in self.mass_excel.items():
            # Выписка из лицевого счета АДБ файл «VXXXXX. xls»
            v_postup_value = 0
            v_vozvrat_value = 0
            v_zachet_value = 0
            length = 0
            try:
                length = len(pd.read_excel(self.file_path + files['V'], sheet_name=None))
            except:
                length = 0
            try:
                v_workbook = pd.read_excel(self.file_path + files['V'], sheet_name=length-1)
                values = v_workbook.values

                for i in range(len(values)):
                    for j in range(len(values[i])):
                        if str(values[i][j]).__contains__("Поступления"):
                            v_postup_col = j
                        if str(values[i][j]).__contains__("Возвраты"):
                            v_vozvrat_col = j
                        if str(values[i][j]).__contains__("Зачеты"):
                            v_zachet_col = j
                        if str(values[i][j]).__contains__("на конец дня"):
                            v_postup_value = float(str(values[i][v_postup_col]).replace(' ','').replace(',','.'))
                            v_vozvrat_value = float(str(values[i][v_vozvrat_col]).replace(' ','').replace(',','.'))
                            v_zachet_value = float(str(values[i][v_zachet_col]).replace(' ','').replace(',','.'))
            except:
                v_workbook = pd.read_html(self.file_path + files['V'])
                frame = v_workbook[0]
                values = frame.values

                for i in range(len(values)):
                    for j in range(len(values[i])):
                        if str(values[i][j]).__contains__("Поступления"):
                            v_postup_col = j
                        if str(values[i][j]).__contains__("Возвраты"):
                            v_vozvrat_col = j
                        if str(values[i][j]).__contains__("Зачеты"):
                            v_zachet_col = j
                        if str(values[i][j]).__contains__("на конец дня"):
                            v_postup_value = float(str(values[i][v_postup_col]).replace(' ','').replace(',','.'))
                            v_vozvrat_value = float(str(values[i][v_vozvrat_col]).replace(' ','').replace(',','.'))
                            v_zachet_value = float(str(values[i][v_zachet_col]).replace(' ','').replace(',','.'))


            print("v_postup: "+str(v_postup_value))
            print("v_vozvrat: "+str(v_vozvrat_value))
            print("v_zachet: "+str(v_zachet_value))

            # Отчет о состоянии лицевого счета АДБ файл «OXXXXX. xls»
            o_postup_value = 0
            o_vozvrat_value = 0
            o_zachet_value = 0
            o_itogo_value = 0
            length = 0
            try:
                length = len(pd.read_excel(self.file_path + files['O'], sheet_name=None))
            except:
                length = 0
            try:
                # Если формат excel то выполнется корректно
                if(length > 2):
                    # Костыль ебаный я в ахуе просто тотальном. В файле может быть и 3 листа
                    length = 2   
                if(length == 1):
                    length = 1
                o_workbook = pd.read_excel(self.file_path + files['O'], sheet_name=length-1)
                values = o_workbook.values

                for i in range(len(values)):
                    _break = False
                    for j in range(len(values[i])):
                        if str(values[i][j]).__contains__("Поступления"):
                            o_postup_col = j
                        if str(values[i][j]).__contains__("Возвраты"):
                            o_vozvrat_col = j
                        if str(values[i][j]).__contains__("Зачеты"):
                            o_zachet_col = j
                        if str(values[i][j]).__contains__("Дата"):
                            o_date_row = i
                            o_date_col = j
                        if str(values[i][j]) == "Итого" or str(values[i][j]).__contains__("(гр.3-гр.4+гр.5)"):
                            o_itogo_col = j
                        if str(values[i][j]) == "Итого:":
                            o_postup_value = float(str(values[i][o_postup_col]).replace(" ", "").replace(",", "."))
                            o_vozvrat_value = float(str(values[i][o_vozvrat_col]).replace(' ','').replace(',','.').replace(u'\xa0', ''))
                            o_zachet_value = float(str(values[i][o_zachet_col]).replace(' ','').replace(',','.').replace(u'\xa0', ''))
                            o_itogo_value = float(str(values[i][o_itogo_col]).replace(' ','').replace(',','.').replace(u'\xa0', ''))
                            _break = True

                    if _break:
                        break
                j = o_date_col
                o_date = ""
                while j < len(values[o_date_row]):
                    if str(values[o_date_row][j]) == 'nan':
                        j+=1
                    else:
                        o_date = str(values[o_date_row][j])
                        j+=1


                    
            except ValueError as ex:
                print("Ошибка при считывании файла O")
                # Если это скрытый html под excel, то считываем иначе
                o_workbook = pd.read_html(self.file_path + files['O'])
                values = o_workbook[0].values

                for i in range(len(values)):
                    _break = False
                    for j in range(len(values[i])):
                        if str(values[i][j]).__contains__("Поступления"):
                            o_postup_col = j
                        if str(values[i][j]).__contains__("Возвраты"):
                            o_vozvrat_col = j
                        if str(values[i][j]).__contains__("Зачеты"):
                            o_zachet_col = j
                        if str(values[i][j]).__contains__("Дата"):
                            o_date_row = i
                            o_date_col = j
                        if str(values[i][j]) == "Итого" or str(values[i][j]).__contains__("(гр.3-гр.4+гр.5)"):
                            o_itogo_col = j
                        if str(values[i][j]) == "Итого:":
                            o_postup_value = float(str(values[i][o_postup_col]).replace(' ','').replace(',','.'))
                            o_vozvrat_value = float(str(values[i][o_vozvrat_col]).replace(' ','').replace(',','.'))
                            o_zachet_value = float(str(values[i][o_zachet_col]).replace(' ','').replace(',','.'))
                            o_itogo_value = float(str(values[i][o_itogo_col]).replace(' ','').replace(',','.'))
                            _break = True

                    if _break:
                        break

                j = o_date_col
                o_date = ""
                while j < len(values[o_date_row]):
                    if str(values[o_date_row][j]) == 'nan':
                        j+=1
                    else:
                        o_date = str(values[o_date_row][j])
                        j+=1
            
            print("o_postup: "+str(o_postup_value))
            print("o_vozvrat: "+str(o_vozvrat_value))
            print("o_zachet: "+str(o_zachet_value))
            print("o_itogo: "+str(o_itogo_value))
            

            # Справка о перечислении поступлений в бюджеты для АДБ файл «CXXXXX.xls»
            c_workbook = pd.read_excel(self.file_path + files['C'], sheet_name=1)
            values = c_workbook.values

            c_perechislen_value = 0
            c_ostatok_value = 0
            c_postup_vsego_value = 0

            for i in range(len(values)):
                if str(values[i][2]) == "Всего по разделам I и II":
                    c_perechislen_value = float(str(values[i][7]).replace(' ','').replace(',','.'))


            c_workbook = pd.read_excel(self.file_path + files['C'], sheet_name=2)
            values = c_workbook.values

            for i in range(len(values)):
                if str(values[i][2]) == "Всего по разделам I и II":
                    c_ostatok_value = float(str(values[i][14]).replace(' ','').replace(',','.'))

            c_workbook = pd.read_excel(self.file_path + files['C'], sheet_name=3)
            values = c_workbook.values

            for i in range(len(values)):
                if str(values[i][2]) == "Всего по разделу III":
                    c_postup_vsego_value = float(str(values[i][3]).replace(' ','').replace(',','.'))

            print("c_perechisl: "+str(c_perechislen_value))
            print("c_ostatok: "+str(c_ostatok_value))
            print("c_postup_vsego: "+str(c_postup_vsego_value))

            temp = {"ubp": "", "date": o_date, "results": []}
            self.protokol.emit("<p>Результат сверки для УБП {}".format(ubp))
            temp["ubp"] = "Результат сверки для УБП {}".format(ubp)
            # self.excel_protokol.emit("Результат сверки для УБП {}".format(ubp), "")

            if "{:.2f}".format(v_postup_value) == "{:.2f}".format(o_postup_value):
                self.protokol.emit("&nbsp;&nbsp;&nbsp;&nbsp;Поступления в выписке и отчете <font color='green'>Равны</font>")
                temp["results"].append({'descr': "              Поступления в выписке и отчете", "result": "Равны"})
                # self.excel_protokol.emit("              Поступления в выписке и отчете", "Равны")
            else:
                self.protokol.emit("&nbsp;&nbsp;&nbsp;&nbsp;Поступления в выписке и отчете <font color='red'>Различаются на {:.2f}</font>".format(v_postup_value-o_postup_value))
                temp["results"].append({'descr': "              Поступления в выписке и отчете", "result": "Различаются на {:.2f}".format(v_postup_value-o_postup_value)})
                # self.excel_protokol.emit("              Поступления в выписке и отчете", "Различаются на {:.2f}".format(v_postup_value-o_postup_value))

            if "{:.2f}".format(v_vozvrat_value) == "{:.2f}".format(o_vozvrat_value):
                self.protokol.emit("&nbsp;&nbsp;&nbsp;&nbsp;Возвраты в выписке и отчете <font color='green'>Равны</font>")
                temp["results"].append({'descr': "              Возвраты в выписке и отчете", "result": "Равны"})
                # self.excel_protokol.emit("              Возвраты в выписке и отчете", "Равны")
            else:
                self.protokol.emit("&nbsp;&nbsp;&nbsp;&nbsp;Возвраты в выписке и отчете <font color='red'>Различаются на {:.2f}</font>".format(v_vozvrat_value-o_vozvrat_value))
                temp["results"].append({'descr': "              Возвраты в выписке и отчете", "result": "Различаются на {:.2f}".format(v_vozvrat_value-o_vozvrat_value)})
                # self.excel_protokol.emit("              Возвраты в выписке и отчете", "Различаются на {:.2f}".format(v_vozvrat_value-o_vozvrat_value))
            
            if "{:.2f}".format(v_zachet_value) == "{:.2f}".format(o_zachet_value):
                self.protokol.emit("&nbsp;&nbsp;&nbsp;&nbsp;Зачеты в выписке и отчете <font color='green'>Равны</font>")
                temp["results"].append({'descr': "              Зачеты в выписке и отчете", "result": "Равны"})
                # self.excel_protokol.emit("              Зачеты в выписке и отчете", "Равны")
            else:
                self.protokol.emit("&nbsp;&nbsp;&nbsp;&nbsp;Зачеты в выписке и отчете <font color='red'>Различаются на {:.2f}</font>".format(v_zachet_value-o_zachet_value))
                temp["results"].append({'descr': "              Зачеты в выписке и отчете", "result": "Различаются на {:.2f}".format(v_zachet_value-o_zachet_value)})
                # self.excel_protokol.emit("              Зачеты в выписке и отчете", "Различаются на {:.2f}".format(v_zachet_value-o_zachet_value))

            if "{:.2f}".format(o_itogo_value) == "{:.2f}".format(c_perechislen_value + c_ostatok_value + c_postup_vsego_value):
                self.protokol.emit("&nbsp;&nbsp;&nbsp;&nbsp;Итого в отчете и справке <font color='green'>Равны</font></p>")
                temp["results"].append({'descr': "              Итого в отчете и справке", "result": "Равны"})
                # self.excel_protokol.emit("              Итого в отчете и справке", "Равны")
            else:
                self.protokol.emit("&nbsp;&nbsp;&nbsp;&nbsp;Итого в отчете и справке <font color='red'>Различаются на {:.2f}</font></p>".format((c_perechislen_value + c_ostatok_value + c_postup_vsego_value)-o_itogo_value))
                temp["results"].append({'descr': "              Итого в отчете и справке", "result": "Различаются на {:.2f}".format((c_perechislen_value + c_ostatok_value + c_postup_vsego_value)-o_itogo_value)})
                # self.excel_protokol.emit("              Итого в отчете и справке", "Различаются на {:.2f}".format((c_perechislen_value + c_ostatok_value + c_postup_vsego_value)-o_itogo_value))
            
            self.excel_protokol.emit(temp)

        self.pb_change_visible.emit()