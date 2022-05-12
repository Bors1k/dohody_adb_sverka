import pandas as pd


v_workbook = pd.read_excel("C:\\Users\\ufk4800_mnovikov\\Desktop\\Юля\\O01982.xls", sheet_name=0)
# print(len(v_workbook))
values = v_workbook.values

o_postup_value = 0.00
o_vozvrat_value = 0.00
o_zachet_value = 0.00
o_itogo_value = 0.00

for i in range(len(values)):
    _break = False
    for j in range(len(values[i])):
        if str(values[i][j]).__contains__("Поступления"):
            o_postup_col = j
        if str(values[i][j]).__contains__("Возвраты"):
            o_vozvrat_col = j
        if str(values[i][j]).__contains__("Зачеты"):
            o_zachet_col = j
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


print(o_postup_value)
print(o_vozvrat_value)
print(o_zachet_value)
print(o_itogo_value)