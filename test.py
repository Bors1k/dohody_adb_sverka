import pandas as pd


o_workbook = pd.read_html("C:\\Users\\ufk4800_mnovikov\\Desktop\\O00001.xls")
values = o_workbook[0].values

o_postup_value = 0
o_vozvrat_value = 0
o_zachet_value = 0
o_itogo_value = 0

for i in range(len(values)):
    for j in range(len(values[i])):
        if str(values[i][j]).__contains__("Поступления"):
            o_postup_col = j
        if str(values[i][j]).__contains__("Возвраты"):
            o_vozvrat_col = j
        if str(values[i][j]).__contains__("Зачеты"):
            o_zachet_col = j
        if str(values[i][j]) == "Итого":
            o_itogo_col = j
        if str(values[i][j]) == "Итого:":
            o_postup_value = float(values[i][o_postup_col].replace(' ','').replace(',','.'))
            o_vozvrat_value = float(values[i][o_vozvrat_col].replace(' ','').replace(',','.'))
            o_zachet_value = float(values[i][o_zachet_col].replace(' ','').replace(',','.'))
            o_itogo_value = float(values[i][o_itogo_col].replace(' ','').replace(',','.'))

print(o_postup_value)
print(o_vozvrat_value)
print(o_zachet_value)
print(o_itogo_value)