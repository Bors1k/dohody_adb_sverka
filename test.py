import pandas as pd


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