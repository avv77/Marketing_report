import os
import comtypes.client
from docxtpl import DocxTemplate, InlineImage
from chart1 import chart1, chart2, chart3
from custom_selection import import_base_processing, variation, variation_stoim, variation_stoim_ros
from file_import import dynamics, ros_stoim_table, transformation_value, dynamics_value_year, cost_dol
from pivot_table_country import pivot1_1, pivot1_2, pivot1_3, pivot2_1, pivot2_2, pivot_table_country_year, pivotfo, \
    pivot_fo_1, pivot_table_fo_reg, pivot_country_cost_excel, pivot_country_cost_word, pivot_table_country_dol, \
    pivot_country_cost_rub
from processing_files_exel import tnved_number
from production import production_base, production_delete_value, production_dict_value_sum, \
    production_dict_dynamics, pivot_pro_1, pivot_pro_2, pivot_table_fo_year
from settings import year_now, year_last, products1, year2013, year2014, year2015, year2016, year2017, year2018, \
    year2019, year2020, year2021, file_report_pattern, file_chart_png2, list_year, path_exel2, file_chart_png3, \
    file_chart_png4, file_report_pattern_final, production, file_chart_png5, index_production, index_import, \
    index_export, file_chart_png6, file_chart_png7, file_chart_png8, file_chart_png9, \
    file_chart_png10, file_chart_png11, path_produst, year_int_2013, year_int_2014, year_int_2015, year_int_2016, \
    year_int_2017, year_int_2018, year_int_2019, year_int_2020, year_int_2021, dimension_ros_stoim, dimension_netto
from docx.shared import Cm
from settings import file_chart_png
import pandas as pd
import sys
from PyPDF2 import PdfReader, PdfWriter


tnved_code_name = tnved_number(path_produst, production)

tnved1 = tnved_code_name[4]
tnved2 = tnved_code_name[5]
tnved3 = tnved_code_name[6]
name_tnved1 = tnved_code_name[0]
name_tnved2 = tnved_code_name[1]
name_tnved3 = tnved_code_name[2]

import_base = import_base_processing('ИМ', path_produst, production)

# глава 5.1
dict_cost_dol = cost_dol()
variation_cost_dol = variation(dict_cost_dol)
doc = DocxTemplate(file_report_pattern)
chart1(dict_cost_dol, '$/кг', file_chart_png8)
image6 = InlineImage(doc, file_chart_png8, Cm(16.5))
dynamics_cost = dynamics(variation_cost_dol[1][0], variation_cost_dol[1][len(variation_cost_dol[1]) - 1])
if variation_cost_dol[1][0] != 0:
    dynamics_cost_val = round(variation_cost_dol[1][len(variation_cost_dol[1]) - 1] / variation_cost_dol[1][0] * 100 -
                              100, 1)
else:
    dynamics_cost_val = 'многократное количество '
min_year_cost = min(dict_cost_dol, key=dict_cost_dol.get)
value_min_year_cost = dict_cost_dol[min_year_cost]
max_year_cost = max(dict_cost_dol, key=dict_cost_dol.get)
value_max_year_cost = dict_cost_dol[max_year_cost]
ros_stoim_cost = ros_stoim_table(dict_cost_dol, dimension_netto)
variation_cost_rub = variation(ros_stoim_cost)
chart1(ros_stoim_cost, 'руб./кг', file_chart_png9)
image7 = InlineImage(doc, file_chart_png9, Cm(16.5))
dynamics_cost_rub = dynamics(variation_cost_rub[1][0], variation_cost_rub[1][len(variation_cost_rub[1]) - 1])

if variation_cost_rub[1][0] != 0:
    dynamics_cost_val_rub = round(variation_cost_rub[1][len(variation_cost_rub[1]) - 1] / variation_cost_rub[1][0] *
                                  100 - 100, 1)
else:
    dynamics_cost_val_rub = 'многократное количество '

min_year_cost_rub = min(ros_stoim_cost, key=ros_stoim_cost.get)
value_min_year_cost_rub = ros_stoim_cost[min_year_cost_rub]
max_year_cost_rub = max(ros_stoim_cost, key=ros_stoim_cost.get)
value_max_year_cost_rub = ros_stoim_cost[max_year_cost_rub]

# глава 5.2
pivot1_1(list_year[len(list_year) - 1])
pivot_country_cost_excel()
table_country_dol = pivot_country_cost_word()
# заполняем сводную таблицу по странам цене в долларах с 2013 г.
table_country_dol_din = pivot_table_country_dol()
table_country_rub = pivot_country_cost_rub()

chart1(import_base[4], import_base[3], file_chart_png)
image = InlineImage(doc, file_chart_png, Cm(16.5))
variation_value = variation(import_base[4])
if variation_value[1][0] != 0:
    dynamics_netto = round(variation_value[1][len(variation_value[1]) - 1] / variation_value[1][0] * 100 - 100, 1)
else:
    dynamics_netto = 'многократное количество '

max_year = max(import_base[4], key=import_base[4].get)
value_max_year = import_base[4][max_year]
min_year = min(import_base[4], key=import_base[4].get)
value_min_year = import_base[4][min_year]
variation_value_stoim = variation_stoim(import_base[8])
if variation_value_stoim[1][0] != 0:
    dynamics_stoim = round(variation_value_stoim[1][len(variation_value_stoim[1]) - 1] /
                       variation_value_stoim[1][0] * 100 - 100, 1)
else:
    dynamics_stoim = 'многократное количество '
dynamics1 = dynamics(variation_value[1][0], variation_value[1][len(variation_value[1]) - 1])
max_year_stoim = max(import_base[8], key=import_base[8].get)
value_max_year_stoim = import_base[8][max_year_stoim]
min_year_stoim = min(import_base[8], key=import_base[8].get)
value_min_year_stoim = import_base[8][min_year_stoim]
dynamics2 = dynamics(variation_value_stoim[1][0], variation_value_stoim[1][len(variation_value_stoim[1]) - 1])
# заполняем таблицу - стоимость в рублях
ros_stoim = ros_stoim_table(import_base[8], dimension_ros_stoim)
# заполняем таблицу - динамика стоимости в рублях
variation_value_stoim_ros = variation_stoim_ros(ros_stoim)
# заполняем текстовку после таблицы
dynamics3 = dynamics(variation_value_stoim_ros[1][0], variation_value_stoim_ros[1][len(variation_value_stoim_ros[1]) -
                                                                                   1])
if variation_value_stoim_ros[1][0] != 0:
    dynamics_stoim_ros = round(variation_value_stoim_ros[1][len(variation_value_stoim[1]) - 1] /
                           variation_value_stoim_ros[1][0] * 100 - 100, 1)
else:
    dynamics_stoim_ros = 'многократное количество '
max_year_stoim_ros = max(ros_stoim, key=ros_stoim.get)
value_max_year_stoim_ros = ros_stoim[max_year_stoim_ros]
min_year_stoim_ros = min(ros_stoim, key=ros_stoim.get)
value_min_year_stoim_ros = ros_stoim[min_year_stoim_ros]

# заполняем сводную таблицу по странам в кг

pivot1_2()
pivot3 = pivot1_3()
table_contents = pivot3[0]
country1 = table_contents[0]['Страны']
country2 = table_contents[1]['Страны']
country3 = table_contents[2]['Страны']
country4 = table_contents[3]['Страны']
country_part = round(pivot3[1], 1)

# делаем и вставляем круговую диаграмму по странам
chart2(pivot3[2], file_chart_png2)
image2 = InlineImage(doc, file_chart_png2, Cm(16.5))

# заполняем таблицу с 4-мя крупнейшими странами
country_larg1 = transformation_value(pivot3[2][0]['NETTO'])
country_larg1_tr = pivot3[2][0]['NETTO']
country_larg2 = transformation_value(pivot3[2][1]['NETTO'])
country_larg2_tr = pivot3[2][1]['NETTO']
country_larg3 = transformation_value(pivot3[2][2]['NETTO'])
country_larg3_tr = pivot3[2][2]['NETTO']
country_larg4 = transformation_value(pivot3[2][3]['NETTO'])
country_larg4_tr = pivot3[2][3]['NETTO']

df_itog1 = pd.read_excel(path_exel2)
itog_all_ = round(df_itog1["NETTO"].sum())
itog_all = transformation_value(itog_all_)

country_other1 = round(itog_all_ - pivot3[2][0]['NETTO'] - pivot3[2][1]['NETTO'] - pivot3[2][2]['NETTO'] -
                       pivot3[2][3]['NETTO'])
country_other1 = transformation_value(country_other1)

pivot1_1(list_year[len(list_year) - 2])
pivot1_2()
pivot3_2 = pivot1_3()

country_larg1_1 = 0
country_larg1_2 = 0
country_larg1_3 = 0
country_larg1_4 = 0

country_larg1_1_tr = 0
country_larg1_2_tr = 0
country_larg1_3_tr = 0
country_larg1_4_tr = 0


for i_dict in range(len(pivot3_2[2]) - 1):
    if country1 == pivot3_2[2][i_dict]['Страны']:
        country_larg1_1 = transformation_value(pivot3_2[2][i_dict]['NETTO'])
        country_larg1_1_tr = pivot3_2[2][i_dict]['NETTO']
for i_dict in range(len(pivot3_2[2]) - 1):
    if country2 == pivot3_2[2][i_dict]['Страны']:
        country_larg1_2 = transformation_value(pivot3_2[2][i_dict]['NETTO'])
        country_larg1_2_tr = pivot3_2[2][i_dict]['NETTO']
for i_dict in range(len(pivot3_2[2]) - 1):
    if country3 == pivot3_2[2][i_dict]['Страны']:
        country_larg1_3 = transformation_value(pivot3_2[2][i_dict]['NETTO'])
        country_larg1_3_tr = pivot3_2[2][i_dict]['NETTO']
for i_dict in range(len(pivot3_2[2]) - 1):
    if country4 == pivot3_2[2][i_dict]['Страны']:
        country_larg1_4 = transformation_value(pivot3_2[2][i_dict]['NETTO'])
        country_larg1_4_tr = pivot3_2[2][i_dict]['NETTO']


df_itog2 = pd.read_excel(path_exel2)
itog_all_2_ = round(df_itog2["NETTO"].sum())
itog_all_2 = transformation_value(itog_all_2_)

len_pivot3_2 = len(pivot3_2[2])

if len_pivot3_2 == 2:
    country_other2 = round(itog_all_2_ - pivot3_2[2][0]['NETTO'])
    country_other2 = transformation_value(country_other2)
elif len_pivot3_2 == 3:
    country_other2 = round(itog_all_2_ - pivot3_2[2][0]['NETTO'] - pivot3_2[2][1]['NETTO'])
    country_other2 = transformation_value(country_other2)

elif len_pivot3_2 == 4:
    country_other2 = round(itog_all_2_ - pivot3_2[2][0]['NETTO'] - pivot3_2[2][1]['NETTO'] - pivot3_2[2][2]['NETTO'])
    country_other2 = transformation_value(country_other2)

else:
    country_other2 = round(itog_all_2_ - pivot3_2[2][0]['NETTO'] - pivot3_2[2][1]['NETTO'] - pivot3_2[2][2]['NETTO'] -
                           pivot3_2[2][3]['NETTO'])
    country_other2 = transformation_value(country_other2)

if len(pivot3[2]) == 2:
    if country_larg1_1_tr != 0:
        dynamics_now_last1 = round(country_larg1_tr / country_larg1_1_tr * 100 - 100, 1)
    else:
        dynamics_now_last1 = ''
    dynamics_now_last2 = ''
    dynamics_now_last3 = ''
    dynamics_now_last4 = ''
    dynamics_now_last5 = round(itog_all_ / itog_all_2_ * 100 - 100, 1)
    if country_larg1_1_tr != 0:
        dynamics_text1 = dynamics(country_larg1_1_tr, country_larg1_tr)
    else:
        dynamics_text1 = ''
    dynamics_text2 = ''
    dynamics_text3 = ''
    dynamics_text4 = ''

elif len(pivot3[2]) == 3:
    if country_larg1_1_tr != 0:
        dynamics_now_last1 = round(country_larg1_tr / country_larg1_1_tr * 100 - 100, 1)
    else:
        dynamics_now_last1 = ''
    if country_larg1_2_tr != 0:
        dynamics_now_last2 = round(country_larg2_tr / country_larg1_2_tr * 100 - 100, 1)
    else:
        dynamics_now_last2 = ''
    dynamics_now_last3 = ''
    dynamics_now_last4 = ''
    dynamics_now_last5 = round(itog_all_ / itog_all_2_ * 100 - 100, 1)

    if country_larg1_1_tr != 0:
        dynamics_text1 = dynamics(country_larg1_1_tr, country_larg1_tr)
    else:
        dynamics_text1 = ''
    if country_larg2_tr != 0:
        dynamics_text2 = dynamics(country_larg1_2_tr, country_larg2_tr)
    else:
        dynamics_text2 = ''

    dynamics_text3 = ''
    dynamics_text4 = ''

elif len(pivot3[2]) == 4:
    if country_larg1_1_tr != 0:
        dynamics_now_last1 = round(country_larg1_tr / country_larg1_1_tr * 100 - 100, 1)
    else:
        dynamics_now_last1 = ''
    if country_larg1_2_tr != 0:
        dynamics_now_last2 = round(country_larg2_tr / country_larg1_2_tr * 100 - 100, 1)
    else:
        dynamics_now_last2 = ''
    if country_larg1_3_tr != 0:
        dynamics_now_last3 = round(country_larg3_tr / country_larg1_3_tr * 100 - 100, 1)
    else:
        dynamics_now_last3 = ''
    dynamics_now_last4 = ''
    dynamics_now_last5 = round(itog_all_ / itog_all_2_ * 100 - 100, 1)

    if country_larg1_1_tr != 0:
        dynamics_text1 = dynamics(country_larg1_1_tr, country_larg1_tr)
    else:
        dynamics_text1 = ''
    if country_larg2_tr != 0:
        dynamics_text2 = dynamics(country_larg1_2_tr, country_larg2_tr)
    else:
        dynamics_text2 = ''
    if country_larg1_3_tr != 0:
        dynamics_text3 = dynamics(country_larg1_3_tr, country_larg3_tr)
    else:
        dynamics_text3 = ''
    dynamics_text4 = ''

else:
    if country_larg1_1_tr != 0:
        dynamics_now_last1 = round(country_larg1_tr / country_larg1_1_tr * 100 - 100, 1)
    else:
        dynamics_now_last1 = ''
    if country_larg1_2_tr != 0:
        dynamics_now_last2 = round(country_larg2_tr / country_larg1_2_tr * 100 - 100, 1)
    else:
        dynamics_now_last2 = ''
    if country_larg1_3_tr != 0:
        dynamics_now_last3 = round(country_larg3_tr / country_larg1_3_tr * 100 - 100, 1)
    else:
        dynamics_now_last3 = ''
    if country_larg1_4_tr != 0:
        dynamics_now_last4 = round(country_larg4_tr / country_larg1_4_tr * 100 - 100, 1)
    else:
        dynamics_now_last4 = ''

    dynamics_now_last5 = round(itog_all_ / itog_all_2_ * 100 - 100, 1)

    if country_larg1_1_tr != 0:
        dynamics_text1 = dynamics(country_larg1_1_tr, country_larg1_tr)
    else:
        dynamics_text1 = ''
    if country_larg2_tr != 0:
        dynamics_text2 = dynamics(country_larg1_2_tr, country_larg2_tr)
    else:
        dynamics_text2 = ''
    if country_larg1_3_tr != 0:
        dynamics_text3 = dynamics(country_larg1_3_tr, country_larg3_tr)
    else:
        dynamics_text3 = ''
    if country_larg1_4_tr != 0:
        dynamics_text4 = dynamics(country_larg1_4_tr, country_larg4_tr)
    else:
        dynamics_text4 = ''

year_contrast1 = list_year[len(list_year) - 2]
year_contrast2 = list_year[len(list_year) - 1]

# заполняем таблицу со странами по стоимости
pivot1_1(list_year[len(list_year) - 1])
pivot2_1()
pivot4 = pivot2_2()
table_contents2 = pivot4[0]
country1_1 = table_contents2[0]['Страны']

if table_contents2[1]['Страны'] == 'Итого':
    country2_1 = ''
    country3_1 = ''
    country4_1 = ''
    country_part_2 = 100.1

elif table_contents2[2]['Страны'] == 'Итого':
    country2_1 = table_contents2[1]['Страны']
    country3_1 = ''
    country4_1 = ''
    country_part_2 = 100.0

elif table_contents2[3]['Страны'] == 'Итого':
    country2_1 = table_contents2[1]['Страны']
    country3_1 = table_contents2[2]['Страны']
    country4_1 = ''
    country_part_2 = 100.0

elif table_contents2[4]['Страны'] == 'Итого':
    country2_1 = table_contents2[1]['Страны']
    country3_1 = table_contents2[2]['Страны']
    country4_1 = table_contents2[3]['Страны']
    country_part_2 = 100.0

else:
    country2_1 = table_contents2[1]['Страны']
    country3_1 = table_contents2[2]['Страны']
    country4_1 = table_contents2[3]['Страны']
    country_part_2 = pivot4[1]

# заполняем сводную таблицу по странам по годам с 2013 г.
table_contents3 = pivot_table_country_year()

# заполняем таблицу с федеральными округами в кг
pivot_fo_1('ФО', 'NETTO', r'exel\report_fo.xlsx')
pivot_fo = pivotfo(r'exel\report_fo.xlsx', 'ФО', 'NETTO')
table_contents_fo = pivot_fo[0]
fo1 = table_contents_fo[0]['ФО']
if table_contents_fo[1]['ФО'] == 'Итого':
    fo2 = ''
    fo3 = ''
    fo4 = ''
    fo_sum_4 = 100.0
elif table_contents_fo[2]['ФО'] == 'Итого':
    fo2 = table_contents_fo[1]['ФО']
    fo3 = ''
    fo4 = ''
    fo_sum_4 = 100.0
elif table_contents_fo[3]['ФО'] == 'Итого':
    fo2 = table_contents_fo[1]['ФО']
    fo3 = table_contents_fo[2]['ФО']
    fo4 = ''
    fo_sum_4 = 100.0

elif table_contents_fo[4]['ФО'] == 'Итого':
    fo2 = table_contents_fo[1]['ФО']
    fo3 = table_contents_fo[2]['ФО']
    fo4 = table_contents_fo[3]['ФО']
    fo_sum_4 = 100.0

else:
    fo2 = table_contents_fo[1]['ФО']
    fo3 = table_contents_fo[2]['ФО']
    fo4 = table_contents_fo[3]['ФО']
    fo_sum_4 = round(pivot_fo[1], 1)

# заполняем таблицу с регионами в кг
pivot_fo_1('Регионы', 'NETTO', r'exel\report_reg.xlsx')
pivot_reg = pivotfo(r'exel\report_reg.xlsx', 'Регионы', 'NETTO')
table_contents_reg = pivot_reg[0]
reg1 = table_contents_reg[0]['Регионы']
if table_contents_reg[1]['Регионы'] == 'Итого':
    reg2 = ''
    reg3 = ''
    reg4 = ''
    reg_sum_4 = 100.0

elif table_contents_reg[2]['Регионы'] == 'Итого':
    reg2 = table_contents_reg[1]['Регионы']
    reg3 = ''
    reg4 = ''
    reg_sum_4 = 100.0

elif table_contents_reg[3]['Регионы'] == 'Итого':
    reg2 = table_contents_reg[1]['Регионы']
    reg3 = table_contents_reg[2]['Регионы']
    reg4 = ''
    reg_sum_4 = 100.0

elif table_contents_reg[4]['Регионы'] == 'Итого':
    reg2 = table_contents_reg[1]['Регионы']
    reg3 = table_contents_reg[2]['Регионы']
    reg4 = table_contents_reg[3]['Регионы']
    reg_sum_4 = 100.0

else:
    reg2 = table_contents_reg[1]['Регионы']
    reg3 = table_contents_reg[2]['Регионы']
    reg4 = table_contents_reg[3]['Регионы']
    reg_sum_4 = round(pivot_reg[1], 1)

# заполняем таблицу по федеральным округам и регионам (4.12)
pivot_fo_reg = pivot_table_fo_reg(r'exel\report_fo_reg.xlsx', ['ФО', 'Регионы'], 'NETTO', 'ФО', 'Регионы', path_exel2)

# ЭКСПОРТ
# заполняем динамику по экспорту
export_base = import_base_processing('ЭК', path_produst, production)

# глава 6.1
value_exp_cost = cost_dol()
variation_cost_dol_exp = variation(value_exp_cost)
chart1(value_exp_cost, '$/кг', file_chart_png10)
image8 = InlineImage(doc, file_chart_png10, Cm(16.5))
dynamics_cost_exp = dynamics(variation_cost_dol_exp[1][0], variation_cost_dol_exp[1][len(variation_cost_dol_exp[1]) -
                                                                                     1])
if variation_cost_dol_exp[1][0] != 0:
    dynamics_cost_val_exp = round(variation_cost_dol_exp[1][len(variation_cost_dol_exp[1]) - 1] /
                                  variation_cost_dol_exp[1][0] * 100 - 100, 1)
else:
    dynamics_cost_val_exp = 'многократное количество '
min_year_cost_exp = min(value_exp_cost, key=value_exp_cost.get)
value_min_year_cost_exp = value_exp_cost[min_year_cost_exp]
max_year_cost_exp = max(value_exp_cost, key=value_exp_cost.get)
value_max_year_cost_exp = value_exp_cost[max_year_cost_exp]
ros_stoim_cost_exp = ros_stoim_table(value_exp_cost, dimension_netto)
variation_cost_rub_exp = variation(ros_stoim_cost_exp)
chart1(ros_stoim_cost_exp, 'руб./кг', file_chart_png11)
image9 = InlineImage(doc, file_chart_png11, Cm(16.5))
dynamics_cost_exp_rub = dynamics(variation_cost_rub_exp[1][0], variation_cost_rub_exp[1][len(variation_cost_rub_exp[1])
                                                                                         - 1])
if variation_cost_rub_exp[1][0] != 0:
    dynamics_cost_val_exp_rub = round(variation_cost_rub_exp[1][len(variation_cost_rub_exp[1]) - 1] /
                                      variation_cost_rub_exp[1][0] * 100 - 100, 1)
else:
    dynamics_cost_val_exp_rub = 'многократное количество '
min_year_cost_exp_rub = min(ros_stoim_cost_exp, key=ros_stoim_cost_exp.get)
value_min_year_cost_exp_rub = ros_stoim_cost_exp[min_year_cost_exp_rub]
max_year_cost_exp_rub = max(ros_stoim_cost_exp, key=ros_stoim_cost_exp.get)
value_max_year_cost_exp_rub = ros_stoim_cost_exp[max_year_cost_exp_rub]

# глава 6.2
pivot1_1(list_year[len(list_year) - 1])
pivot_country_cost_excel()
table_country_dol_ex = pivot_country_cost_word()
table_country_dol_din_ex = pivot_table_country_dol()
table_country_rub_ex = pivot_country_cost_rub()

# диаграмма по динамике экспорта
chart1(export_base[4], export_base[3], file_chart_png3)
image_exp = InlineImage(doc, file_chart_png3, Cm(16.5))

# таблица 2.2 и описание ниже
variation_value_exp = variation(export_base[4])
dynamics1_exp = dynamics(variation_value_exp[1][0], variation_value_exp[1][len(variation_value_exp[1]) - 1])
if variation_value_exp[1][0] != 0:
    dynamics_netto_exp = round(variation_value_exp[1][len(variation_value_exp[1]) - 1] / variation_value_exp[1][0] *
                               100 - 100, 1)
else:
    dynamics_netto_exp = 'многократное количество '
max_year_exp = max(export_base[4], key=export_base[4].get)
value_max_year_exp = export_base[4][max_year_exp]
min_year_exp = min(export_base[4], key=export_base[4].get)
value_min_year_exp = export_base[4][min_year_exp]

# таблица 2.3 и 2.4 и описание ниже
variation_value_stoim_exp = variation_stoim(export_base[8])
dynamics2_exp = dynamics(variation_value_stoim_exp[1][0], variation_value_stoim_exp[1][len(variation_value_stoim_exp[1])
                                                                                       - 1])
if variation_value_stoim_exp[1][0] != 0:
    dynamics_stoim_exp = round(variation_value_stoim_exp[1][len(variation_value_stoim_exp[1]) - 1] /
                               variation_value_stoim_exp[1][0] * 100 - 100, 1)
else:
    dynamics_stoim_exp = 'многократное количество '
max_year_stoim_exp = max(export_base[8], key=export_base[8].get)
value_max_year_stoim_exp = export_base[8][max_year_stoim_exp]
min_year_stoim_exp = min(export_base[8], key=export_base[8].get)
value_min_year_stoim_exp = export_base[8][min_year_stoim_exp]

# таблица 2.5 и 2.6 и описание ниже
ros_stoim_exp = ros_stoim_table(export_base[8], dimension_ros_stoim)
variation_value_stoim_ros_exp = variation_stoim_ros(ros_stoim_exp)
dynamics3_exp = dynamics(variation_value_stoim_ros_exp[1][0], variation_value_stoim_ros_exp[1][
    len(variation_value_stoim_ros_exp[1]) - 1])
if variation_value_stoim_ros_exp[1][0] != 0:
    dynamics_stoim_ros_exp = round(variation_value_stoim_ros_exp[1][len(variation_value_stoim_exp[1]) - 1] /
                                   variation_value_stoim_ros_exp[1][0] * 100 - 100, 1)
else:
    dynamics_stoim_ros_exp = 'многократное количество '
max_year_stoim_ros_exp = max(ros_stoim_exp, key=ros_stoim_exp.get)
value_max_year_stoim_ros_exp = ros_stoim_exp[max_year_stoim_ros_exp]
min_year_stoim_ros_exp = min(ros_stoim_exp, key=ros_stoim_exp.get)
value_min_year_stoim_ros_exp = ros_stoim_exp[min_year_stoim_ros_exp]

# таблица 2.7 география экспорта
pivot1_1(list_year[len(list_year) - 1])
pivot1_2()
pivot3_exp = pivot1_3()
table_contents_exp = pivot3_exp[0]
country1_exp = table_contents_exp[0]['Страны']
if table_contents_exp[1]['Страны'] == 'Итого':
    country2_exp = ''
    country3_exp = ''
    country4_exp = ''
    country_part_exp = 100.0

elif table_contents_exp[2]['Страны'] == 'Итого':
    country2_exp = table_contents_exp[1]['Страны']
    country3_exp = ''
    country4_exp = ''
    country_part_exp = 100.0

elif table_contents_exp[3]['Страны'] == 'Итого':
    country2_exp = table_contents_exp[1]['Страны']
    country3_exp = table_contents_exp[2]['Страны']
    country4_exp = ''
    country_part_exp = 100.0

elif table_contents_exp[4]['Страны'] == 'Итого':
    country2_exp = table_contents_exp[1]['Страны']
    country3_exp = table_contents_exp[2]['Страны']
    country4_exp = table_contents_exp[3]['Страны']
    country_part_exp = 100.0

else:
    country2_exp = table_contents_exp[1]['Страны']
    country3_exp = table_contents_exp[2]['Страны']
    country4_exp = table_contents_exp[3]['Страны']
    country_part_exp = round(pivot3_exp[1], 1)

# диаграмма 2.2 география экспорта
chart2(pivot3_exp[2], file_chart_png4)
image2_exp = InlineImage(doc, file_chart_png4, Cm(16.5))

# таблица 2.8 4 страны экспортера
if len(pivot3_exp[2]) == 2:
    country_larg1_exp = transformation_value(pivot3_exp[2][0]['NETTO'])
    country_larg1_tr_exp = pivot3_exp[2][0]['NETTO']
    country_larg2_exp = 'нет данных'
    country_larg2_tr_exp = 'нет данных'
    country_larg3_exp = 'нет данных'
    country_larg3_tr_exp = 'нет данных'
    country_larg4_exp = 'нет данных'
    country_larg4_tr_exp = 'нет данных'
elif len(pivot3_exp[2]) == 3:
    country_larg1_exp = transformation_value(pivot3_exp[2][0]['NETTO'])
    country_larg1_tr_exp = pivot3_exp[2][0]['NETTO']
    country_larg2_exp = transformation_value(pivot3_exp[2][1]['NETTO'])
    country_larg2_tr_exp = pivot3_exp[2][1]['NETTO']
    country_larg3_exp = 'нет данных'
    country_larg3_tr_exp = 'нет данных'
    country_larg4_exp = 'нет данных'
    country_larg4_tr_exp = 'нет данных'
elif len(pivot3_exp[2]) == 4:
    country_larg1_exp = transformation_value(pivot3_exp[2][0]['NETTO'])
    country_larg1_tr_exp = pivot3_exp[2][0]['NETTO']
    country_larg2_exp = transformation_value(pivot3_exp[2][1]['NETTO'])
    country_larg2_tr_exp = pivot3_exp[2][1]['NETTO']
    country_larg3_exp = transformation_value(pivot3_exp[2][2]['NETTO'])
    country_larg3_tr_exp = pivot3_exp[2][2]['NETTO']
    country_larg4_exp = 'нет данных'
    country_larg4_tr_exp = 'нет данных'
else:
    country_larg1_exp = transformation_value(pivot3_exp[2][0]['NETTO'])
    country_larg1_tr_exp = pivot3_exp[2][0]['NETTO']
    country_larg2_exp = transformation_value(pivot3_exp[2][1]['NETTO'])
    country_larg2_tr_exp = pivot3_exp[2][1]['NETTO']
    country_larg3_exp = transformation_value(pivot3_exp[2][2]['NETTO'])
    country_larg3_tr_exp = pivot3_exp[2][2]['NETTO']
    country_larg4_exp = transformation_value(pivot3_exp[2][3]['NETTO'])
    country_larg4_tr_exp = pivot3_exp[2][3]['NETTO']

df_itog1_exp = pd.read_excel(path_exel2)
itog_all_exp = round(df_itog1_exp["NETTO"].sum())
itog_all_exp1 = transformation_value(itog_all_exp)

if len(pivot3_exp[2]) == 2:
    country_other1_exp = round(itog_all_exp - pivot3_exp[2][0]['NETTO'])
elif len(pivot3_exp[2]) == 3:
    country_other1_exp = round(itog_all_exp - pivot3_exp[2][0]['NETTO'] - pivot3_exp[2][1]['NETTO'])
elif len(pivot3_exp[2]) == 4:
    country_other1_exp = round(itog_all_exp - pivot3_exp[2][0]['NETTO'] - pivot3_exp[2][1]['NETTO'] -
                               pivot3_exp[2][2]['NETTO'])
else:
    country_other1_exp = round(itog_all_exp - pivot3_exp[2][0]['NETTO'] - pivot3_exp[2][1]['NETTO'] -
                               pivot3_exp[2][2]['NETTO'] - pivot3_exp[2][3]['NETTO'])
country_other1_exp = transformation_value(country_other1_exp)

pivot1_1(list_year[len(list_year) - 2])
pivot1_2()
pivot3_2_exp = pivot1_3()

country_larg1_1_exp = 0
country_larg1_2_exp = 0
country_larg1_3_exp = 0
country_larg1_4_exp = 0

country_larg1_1_tr_exp = 0
country_larg1_2_tr_exp = 0
country_larg1_3_tr_exp = 0
country_larg1_4_tr_exp = 0


for i_dict in range(len(pivot3_2_exp[2]) - 1):
    if country1_exp == pivot3_2_exp[2][i_dict]['Страны']:
        country_larg1_1_exp = transformation_value(pivot3_2_exp[2][i_dict]['NETTO'])
        country_larg1_1_tr_exp = pivot3_2_exp[2][i_dict]['NETTO']
for i_dict in range(len(pivot3_2_exp[2]) - 1):
    if country2_exp == pivot3_2_exp[2][i_dict]['Страны']:
        country_larg1_2_exp = transformation_value(pivot3_2_exp[2][i_dict]['NETTO'])
        country_larg1_2_tr_exp = pivot3_2_exp[2][i_dict]['NETTO']
for i_dict in range(len(pivot3_2_exp[2]) - 1):
    if country3_exp == pivot3_2_exp[2][i_dict]['Страны']:
        country_larg1_3_exp = transformation_value(pivot3_2_exp[2][i_dict]['NETTO'])
        country_larg1_3_tr_exp = pivot3_2_exp[2][i_dict]['NETTO']
for i_dict in range(len(pivot3_2_exp[2]) - 1):
    if country4_exp == pivot3_2_exp[2][i_dict]['Страны']:
        country_larg1_4_exp = transformation_value(pivot3_2_exp[2][i_dict]['NETTO'])
        country_larg1_4_tr_exp = pivot3_2_exp[2][i_dict]['NETTO']

df_itog2_exp = pd.read_excel(path_exel2)
itog_all_2_exp = round(df_itog2_exp["NETTO"].sum())
itog_all_2_exp1 = transformation_value(itog_all_2_exp)

country_other2_exp = round(itog_all_2_exp - pivot3_2_exp[2][0]['NETTO'] - pivot3_2_exp[2][1]['NETTO'] - pivot3_2_exp[2][
    2]['NETTO'] - pivot3_2_exp[2][3]['NETTO'])
country_other2_exp = transformation_value(country_other2_exp)

if country_larg1_1_tr_exp != 0:
    dynamics_now_last1_exp = round(country_larg1_tr_exp / country_larg1_1_tr_exp * 100 - 100, 1)
else:
    dynamics_now_last1_exp = 'нет динамики'
if country_larg1_2_tr_exp != 0:
    dynamics_now_last2_exp = round(country_larg2_tr_exp / country_larg1_2_tr_exp * 100 - 100, 1)
else:
    dynamics_now_last2_exp = 'нет динамики'
if country_larg1_3_tr_exp != 0:
    dynamics_now_last3_exp = round(country_larg3_tr_exp / country_larg1_3_tr_exp * 100 - 100, 1)
else:
    dynamics_now_last3_exp = 'нет динамики'
if country_larg1_4_tr_exp != 0:
    dynamics_now_last4_exp = round(country_larg4_tr_exp / country_larg1_4_tr_exp * 100 - 100, 1)
else:
    dynamics_now_last4_exp = 'нет динамики'
if itog_all_2_exp != 0:
    dynamics_now_last5_exp = round(itog_all_exp / itog_all_2_exp * 100 - 100, 1)
else:
    dynamics_now_last5_exp = 'нет динамики'

dynamics_text1_exp = ''
dynamics_text2_exp = ''
dynamics_text3_exp = ''
dynamics_text4_exp = ''

if len(pivot3_exp[2]) == 2:
    dynamics_text1_exp = dynamics(country_larg1_1_tr_exp, country_larg1_tr_exp)
    dynamics_text2_exp = 'нет динамики'
    dynamics_text3_exp = 'нет динамики'
    dynamics_text4_exp = 'нет динамики'
elif len(pivot3_exp[2]) == 3:
    dynamics_text1_exp = dynamics(country_larg1_1_tr_exp, country_larg1_tr_exp)
    dynamics_text2_exp = dynamics(country_larg1_2_tr_exp, country_larg2_tr_exp)
    dynamics_text3_exp = 'нет динамики'
    dynamics_text4_exp = 'нет динамики'
elif len(pivot3_exp[2]) == 4:
    dynamics_text1_exp = dynamics(country_larg1_1_tr_exp, country_larg1_tr_exp)
    dynamics_text2_exp = dynamics(country_larg1_2_tr_exp, country_larg2_tr_exp)
    dynamics_text3_exp = dynamics(country_larg1_3_tr_exp, country_larg3_tr_exp)
    dynamics_text4_exp = 'нет динамики'
else:
    dynamics_text1_exp = dynamics(country_larg1_1_tr_exp, country_larg1_tr_exp)
    dynamics_text2_exp = dynamics(country_larg1_2_tr_exp, country_larg2_tr_exp)
    dynamics_text3_exp = dynamics(country_larg1_3_tr_exp, country_larg3_tr_exp)
    dynamics_text4_exp = dynamics(country_larg1_4_tr_exp, country_larg4_tr_exp)

# заполняем сводную таблицу экспорта по странам по годам с 2013 г.Таблица 2.9
table_contents3_exp = pivot_table_country_year()

# заполняем таблицу экспорта со странами по стоимости. Таблица 2.10
pivot1_1(list_year[len(list_year) - 1])
pivot2_1()
pivot4_exp = pivot2_2()
table_contents2_exp = pivot4_exp[0]
country1_1_exp = table_contents2_exp[0]['Страны']
if table_contents2_exp[1]['Страны'] == 'Итого':
    country2_1_exp = ''
    country3_1_exp = ''
    country4_1_exp = ''
    country_part_2_exp = 100.0

elif table_contents2_exp[2]['Страны'] == 'Итого':
    country2_1_exp = table_contents2_exp[1]['Страны']
    country3_1_exp = ''
    country4_1_exp = ''
    country_part_2_exp = 100.0

elif table_contents2_exp[3]['Страны'] == 'Итого':
    country2_1_exp = table_contents2_exp[1]['Страны']
    country3_1_exp = table_contents2_exp[2]['Страны']
    country4_1_exp = ''
    country_part_2_exp = 100.0

elif table_contents2_exp[4]['Страны'] == 'Итого':
    country2_1_exp = table_contents2_exp[1]['Страны']
    country3_1_exp = table_contents2_exp[2]['Страны']
    country4_1_exp = table_contents2_exp[3]['Страны']
    country_part_2_exp = 100.0

else:
    country2_1_exp = table_contents2_exp[1]['Страны']
    country3_1_exp = table_contents2_exp[2]['Страны']
    country4_1_exp = table_contents2_exp[3]['Страны']
    country_part_2_exp = round(pivot4_exp[1], 1)

# заполняем таблицу экспорта с федеральными округами. Таблица 2.11
pivot_fo_1('ФО', 'NETTO', r'exel\report_fo_exp.xlsx')
pivot_fo_exp = pivotfo(r'exel\report_fo_exp.xlsx', 'ФО', 'NETTO')
table_contents_fo_exp = pivot_fo_exp[0]
fo1_exp = table_contents_fo_exp[0]['ФО']
if table_contents_fo_exp[1]['ФО'] == 'Итого':
    fo2_exp = ''
    fo3_exp = ''
    fo4_exp = ''
    fo_sum_4_exp = 100.0

elif table_contents_fo_exp[2]['ФО'] == 'Итого':
    fo2_exp = table_contents_fo_exp[1]['ФО']
    fo3_exp = ''
    fo4_exp = ''
    fo_sum_4_exp = 100.0

elif table_contents_fo_exp[3]['ФО'] == 'Итого':
    fo2_exp = table_contents_fo_exp[1]['ФО']
    fo3_exp = table_contents_fo_exp[2]['ФО']
    fo4_exp = ''
    fo_sum_4_exp = 100.0

elif table_contents_fo_exp[4]['ФО'] == 'Итого':
    fo2_exp = table_contents_fo_exp[1]['ФО']
    fo3_exp = table_contents_fo_exp[2]['ФО']
    fo4_exp = table_contents_fo_exp[3]['ФО']
    fo_sum_4_exp = 100.0

else:
    fo2_exp = table_contents_fo_exp[1]['ФО']
    fo3_exp = table_contents_fo_exp[2]['ФО']
    fo4_exp = table_contents_fo_exp[3]['ФО']
    fo_sum_4_exp = round(pivot_fo_exp[1], 1)

# заполняем таблицу экспорта с регионами в кг. Таблица 2.12
pivot_fo_1('Регионы', 'NETTO', r'exel\report_reg.xlsx')
pivot_reg_exp = pivotfo(r'exel\report_reg.xlsx', 'Регионы', 'NETTO')
table_contents_reg_exp = pivot_reg_exp[0]
reg1_exp = table_contents_reg_exp[0]['Регионы']
if table_contents_reg_exp[1]['Регионы'] == 'Итого':
    reg2_exp = ''
    reg3_exp = ''
    reg4_exp = ''
    reg_sum_4_exp = 100.0

elif table_contents_reg_exp[2]['Регионы'] == 'Итого':
    reg2_exp = table_contents_reg_exp[1]['Регионы']
    reg3_exp = ''
    reg4_exp = ''
    reg_sum_4_exp = 100.0

elif table_contents_reg_exp[3]['Регионы'] == 'Итого':
    reg2_exp = table_contents_reg_exp[1]['Регионы']
    reg3_exp = table_contents_reg_exp[2]['Регионы']
    reg4_exp = ''
    reg_sum_4_exp = 100.0

elif table_contents_reg_exp[4]['Регионы'] == 'Итого':
    reg2_exp = table_contents_reg_exp[1]['Регионы']
    reg3_exp = table_contents_reg_exp[2]['Регионы']
    reg4_exp = table_contents_reg_exp[3]['Регионы']
    reg_sum_4_exp = 100.0

else:
    reg2_exp = table_contents_reg_exp[1]['Регионы']
    reg3_exp = table_contents_reg_exp[2]['Регионы']
    reg4_exp = table_contents_reg_exp[3]['Регионы']
    reg_sum_4_exp = round(pivot_reg_exp[1], 1)

# заполняем таблицу экспорта по федеральным округам и регионам (2.13)
pivot_fo_reg_exp = pivot_table_fo_reg(r'exel\report_fo_reg_exp.xlsx', ['ФО', 'Регионы'], 'NETTO', 'ФО', 'Регионы',
                                      path_exel2)

# выгружаем из общей базы по производству данные по продукции
production_base(production)
production_delete_value(production)
value_production = production_dict_value_sum()
value_production_m = {}
parameter_production_m = ''
if len(str(value_production[0][int(year_now)])) >= 5:
    for keys, value in value_production[0].items():
        value_production_m[keys] = round(float(value) / 1000, 1)
    if value_production[1] == 'кг':
        parameter_production_m = 'тонн'
    elif value_production[1] == 'тонн':
        parameter_production_m = 'тыс. тонн'
    elif value_production[1] == 'тыс. тонн':
        parameter_production_m = 'млн. тонн'
else:
    value_production_m = value_production[0]
    parameter_production_m = value_production[1]
value_year_now = value_production[0][int(year_now)]
parameter_production = value_production[1]
dynamics_now_last_prod = round(value_year_now / value_production[0][int(year_last)] * 100 - 100, 1)
line_dynamics_year_prod = dynamics_value_year(dynamics_now_last_prod)

# диаграмма по объемам производства (Диаграмма 3.1)
chart1(value_production_m, parameter_production_m, file_chart_png5)
image_pro = InlineImage(doc, file_chart_png5, Cm(16.5))

# динамика объемов производства (Таблица 3.2)
dict_dynamics_year = production_dict_dynamics(value_production[0])
dynamics1_pro = dynamics(value_production[0][2013], value_production[0][int(year_now)])
dynamics_now_last_year_pro = round(value_production[0][int(year_now)] / value_production[0][2013] * 100 - 100, 1)
max_year_pro = max(value_production[0], key=value_production[0].get)
value_max_year_pro = value_production_m[max_year_pro]
min_year_pro = min(value_production[0], key=value_production[0].get)
value_min_year_pro = value_production_m[min_year_pro]

# cтруктура производства по федеральным округам (Таблица 3.3)
pivot_pro_1()
table_contents_fo_pro = pivot_pro_2()
fo_1 = table_contents_fo_pro[0]['Регионы']
fo_dol = table_contents_fo_pro[0]['Доля']

# динамика производства по федеральным округам (Таблица 3.4)
table_contents_fo_din = pivot_table_fo_year()

# объемы производства по годам для расчета потребления
value_production_itog = value_production[0]
value_production_itog_index = {}
for key, value in value_production_itog.items():
    now_value = int(value) / index_production
    value_production_itog_index[key] = round(now_value, 1)
value_production_now_year = round(value_production_itog_index[int(year_now)], 1)
parameter_production_now = parameter_production
if parameter_production == 'тонн' and index_production == 1000:
    parameter_production_now = 'тыс. тонн'
elif parameter_production == 'тонн' and index_production == 1:
    parameter_production_now = 'тонн'
elif parameter_production == 'кг' and index_production == 1000:
    parameter_production_now = 'тонн'
elif parameter_production == 'кг' and index_production == 1:
    parameter_production_now = 'кг'

# объемы потребления
production_list_dict = value_production_itog_index

import_list_dict = {}
for key, value in import_base[4].items():
    import_list_dict[key] = round(float(value) / index_import, 1)
export_list_dict = {}
for key, value in export_base[4].items():
    export_list_dict[key] = round(float(value) / index_export, 1)

consumption_dict = {}
for year in production_list_dict:
    year1 = str(year)
    consumption_dict[year1] = round(float(production_list_dict[year]) + float(import_list_dict[year1]) - float(
        export_list_dict[year1]), 1)
# размерность потребления
dimension_consumption = ''
if index_import == 1:
    dimension_consumption = import_base[3]
elif index_import != 1:
    dimension_consumption = parameter_production

# диаграмма объемы потребления (Диаграмма 4.1)
chart1(consumption_dict, dimension_consumption, file_chart_png6)
image4 = InlineImage(doc, file_chart_png6, Cm(16.5))

# динамика потребления (Таблица 4.2)
variation_consumption = variation(consumption_dict)
dynamics4 = dynamics(variation_consumption[1][0], variation_consumption[1][len(variation_consumption[1]) - 1])
dynamics_netto_con = round(variation_consumption[1][len(variation_consumption[1]) - 1] / variation_consumption[1][0] *
                           100 - 100, 1)
max_year_con = max(consumption_dict, key=consumption_dict.get)
value_max_year_con = consumption_dict[max_year_con]
min_year_con = min(consumption_dict, key=consumption_dict.get)
value_min_year_con = consumption_dict[min_year_con]

# доля импорта и производства
part_im = round(float(import_list_dict[year_now] / consumption_dict[year_now] * 100), 1)
part_pro = round(float(100.0 - part_im), 1)

# делаем и вставляем круговую диаграмму по импорту и отечественному производству (Диаграмма 4.2)
table_name1 = ['импорт', 'производство']
if float(export_list_dict[year_now]) > float(production_list_dict[int(year_now)]):
    print('Экспорт больше производства. Проверить')
    sys.exit()
table_value1 = [round(float(import_list_dict[year_now]), 1), round(float(production_list_dict[int(year_now)]) -
                                                                   float(export_list_dict[year_now]), 1)]
chart3(table_name1, table_value1, file_chart_png7)
image5 = InlineImage(doc, file_chart_png7, Cm(16.5))

# динамика доли импорта (Таблица 4.4)
part_im_dict = {}
for year in import_list_dict.keys():
    part_im_dict[year] = round(float(import_list_dict[year]) / float(consumption_dict[year]) * 100, 1)

value_part_im_dict = part_im_dict
dinamic_im = dynamics(value_part_im_dict['2013'], value_part_im_dict['2021'])

context = {'год_написания': year_now, 'год_сравнения': year_last, 'объем': import_base[0], 'продукция1': products1,
           'значение_динамики': import_base[1], 'направление_динамики': import_base[2], 'размерность': import_base[3],
           'год3': year2013, 'год4': year2014, 'год5': year2015, 'год6': year2016, 'год7': year2017, 'год8': year2018,
           'год9': year2019, 'год10': year2020, 'год11': year2021, 'объем3': import_base[4][year2013], 'объем4':
               import_base[4][year2014], 'объем5': import_base[4][year2015], 'объем6': import_base[4][year2016],
           'объем7': import_base[4][year2017], 'объем8': import_base[4][year2018], 'объем9': import_base[4][year2019],
           'объем10': import_base[4][year2020], 'объем11': import_base[4][year2021], 'image1': image, 'рос1':
               variation_value[0][0], 'рос2': variation_value[0][1], 'рос3': variation_value[0][2], 'рос4':
               variation_value[0][3], 'рос5': variation_value[0][4], 'рос6': variation_value[0][5], 'рос7':
               variation_value[0][6], 'рос8': variation_value[0][7], 'процент1': dynamics_netto, 'год_макс':
               max_year, 'объем_макс': value_max_year,
           'год_мин': min_year, 'объем_мин': value_min_year, 'размерность2': import_base[7], 'дол3':
               import_base[8][year2013], 'дол4': import_base[8][year2014], 'дол5': import_base[8][year2015], 'дол6':
               import_base[8][year2016], 'дол7': import_base[8][year2017], 'дол8': import_base[8][year2018], 'дол9':
               import_base[8][year2019], 'дол10': import_base[8][year2020], 'дол11': import_base[8][year2021], 'дин1':
               variation_value_stoim[0][0], 'дин2': variation_value_stoim[0][1], 'дин3': variation_value_stoim[0][2],
           'дин4': variation_value_stoim[0][3], 'дин5': variation_value_stoim[0][4], 'дин6':
               variation_value_stoim[0][5], 'дин7': variation_value_stoim[0][6], 'дин8': variation_value_stoim[0][7],
           'динамик': dynamics1, 'процент2': dynamics_stoim, 'год_макс2': max_year_stoim, 'объем_макс2':
               value_max_year_stoim, 'год_мин2': min_year_stoim, 'объем_мин2': value_min_year_stoim, 'динамик2':
               dynamics2, 'размерность3': import_base[9], 'руб3': ros_stoim[year2013], 'руб4': ros_stoim[year2014],
           'руб5': ros_stoim[year2015], 'руб6': ros_stoim[year2016], 'руб7': ros_stoim[year2017], 'руб8':
               ros_stoim[year2018], 'руб9': ros_stoim[year2019], 'руб10': ros_stoim[year2020], 'руб11':
               ros_stoim[year2021], 'бол1': variation_value_stoim_ros[0][0], 'бол2': variation_value_stoim_ros[0][1],
           'бол3':
               variation_value_stoim_ros[0][2], 'бол4': variation_value_stoim_ros[0][3], 'бол5':
               variation_value_stoim_ros[0][4], 'бол6': variation_value_stoim_ros[0][5], 'бол7':
               variation_value_stoim_ros[0][6], 'бол8': variation_value_stoim_ros[0][7], 'динамик3': dynamics3,
           'процент3': dynamics_stoim_ros, 'год_макс3': max_year_stoim_ros, 'объем_макс3': value_max_year_stoim_ros,
           'год_мин3': min_year_stoim_ros, 'объем_мин3': value_min_year_stoim_ros, 'table_contents': table_contents,
           'страна1': country1, 'страна2': country2, 'страна3': country3, 'страна4': country4, 'страна_доля':
               country_part, 'table_contents2': table_contents2, 'страна1_дол': country1_1, 'страна2_дол': country2_1,
           'страна3_дол': country3_1, 'страна4_дол': country4_1, 'страна_доля_дол': country_part_2, 'image2': image2,
           'год_ср1': year_contrast1, 'год_ср2': year_contrast2, 'объ_стр2_1': country_larg1, 'объ_стр2_2':
               country_larg2, 'объ_стр2_3': country_larg3, 'объ_стр2_4': country_larg4, 'объ_стр2_5': country_other1,
           'объ_стр2_6': itog_all, 'объ_стр1_1': country_larg1_1, 'объ_стр1_2': country_larg1_2, 'объ_стр1_3':
               country_larg1_3, 'объ_стр1_4': country_larg1_4, 'объ_стр1_5': country_other2, 'объ_стр1_6': itog_all_2,
           'дин_стр1': dynamics_now_last1, 'дин_стр2': dynamics_now_last2, 'дин_стр3':
               dynamics_now_last3, 'дин_стр4': dynamics_now_last4, 'дин_стр5': dynamics_now_last5, 'умел_уменьш1':
               dynamics_text1, 'умел_уменьш2': dynamics_text2, 'умел_уменьш3': dynamics_text3, 'умел_уменьш4':
               dynamics_text4, 'table_contents3': table_contents3, 'table_contents_fo': table_contents_fo, 'фо1': fo1,
           'фо2': fo2, 'фо3': fo3, 'фо4': fo4, 'фо_доля': fo_sum_4, 'table_contents_reg': table_contents_reg, 'рег1':
               reg1, 'рег2': reg2, 'рег3': reg3, 'рег4': reg4, 'рег_доля': reg_sum_4,
           'table_fo_reg': pivot_fo_reg, 'объем_э': export_base[0], 'размерность_э': export_base[3],
           'значение_динамики_э': export_base[1], 'направление_динамики_э': export_base[2],
           'экс3': export_base[4][year2013], 'экс4': export_base[4][year2014], 'экс5': export_base[4][year2015], 'экс6':
               export_base[4][year2016], 'экс7': export_base[4][year2017], 'экс8': export_base[4][year2018], 'экс9':
               export_base[4][year2019], 'экс10': export_base[4][year2020], 'экс11': export_base[4][year2021],
           'image_exp_1': image_exp, 'э1': variation_value_exp[0][0], 'э2': variation_value_exp[0][1], 'э3':
               variation_value_exp[0][2], 'э4': variation_value_exp[0][3], 'э5': variation_value_exp[0][4], 'э6':
               variation_value_exp[0][5], 'э7': variation_value_exp[0][6], 'э8': variation_value_exp[0][7], 'динамик_э':
               dynamics1_exp, 'процент_э_1': dynamics_netto_exp, 'год_макс_э': max_year_exp, 'объем_макс_э':
               value_max_year_exp, 'год_мин_э': min_year_exp, 'объем_мин_э': value_min_year_exp, 'размерность2_э':
               export_base[7], 'дол3_э': export_base[8][year2013], 'дол4_э': export_base[8][year2014], 'дол5_э':
               export_base[8][year2015], 'дол6_э': export_base[8][year2016], 'дол7_э': export_base[8][year2017],
           'дол8_э': export_base[8][year2018], 'дол9_э': export_base[8][year2019], 'дол10_э': export_base[8][year2020],
           'дол11_э': export_base[8][year2021], 'дин1_э': variation_value_stoim_exp[0][0], 'дин2_э':
               variation_value_stoim_exp[0][1], 'дин3_э': variation_value_stoim_exp[0][2], 'дин4_э':
               variation_value_stoim_exp[0][3], 'дин5_э': variation_value_stoim_exp[0][4], 'дин6_э':
               variation_value_stoim_exp[0][5], 'дин7_э': variation_value_stoim_exp[0][6], 'дин8_э':
               variation_value_stoim_exp[0][7], 'динамик2_э': dynamics2_exp, 'процент2_э': dynamics_stoim_exp,
           'год_макс2_э': max_year_stoim_exp, 'объем_макс2_э': value_max_year_stoim_exp, 'год_мин2_э':
               min_year_stoim_exp, 'объем_мин2_э': value_min_year_stoim_exp, 'размерность3_э': export_base[9], 'руб3_э':
               ros_stoim_exp[year2013], 'руб4_э': ros_stoim_exp[year2014], 'руб5_э': ros_stoim_exp[year2015], 'руб6_э':
               ros_stoim_exp[year2016], 'руб7_э': ros_stoim_exp[year2017], 'руб8_э': ros_stoim_exp[year2018], 'руб9_э':
               ros_stoim_exp[year2019], 'руб10_э': ros_stoim_exp[year2020], 'руб11_э': ros_stoim_exp[year2021],
           'бол1_э': variation_value_stoim_ros_exp[0][0], 'бол2_э': variation_value_stoim_ros_exp[0][1], 'бол3_э':
               variation_value_stoim_ros_exp[0][2], 'бол4_э': variation_value_stoim_ros_exp[0][3], 'бол5_э':
               variation_value_stoim_ros_exp[0][4], 'бол6_э': variation_value_stoim_ros_exp[0][5], 'бол7_э':
               variation_value_stoim_ros_exp[0][6], 'бол8_э': variation_value_stoim_ros_exp[0][7], 'динамик3_э':
               dynamics3_exp, 'процент3_э': dynamics_stoim_ros_exp, 'год_макс3_э': max_year_stoim_ros_exp,
           'объем_макс3_э': value_max_year_stoim_ros_exp, 'год_мин3_э': min_year_stoim_ros_exp, 'объем_мин3_э':
               value_min_year_stoim_ros_exp, 'table_contents_exp': table_contents_exp, 'страна1_э': country1_exp,
           'страна2_э': country2_exp, 'страна3_э': country3_exp, 'страна4_э': country4_exp, 'страна_доля_э':
               country_part_exp, 'image2_э': image2_exp, 'объ_стр2_1_э': country_larg1_exp, 'объ_стр2_2_э':
               country_larg2_exp, 'объ_стр2_3_э': country_larg3_exp, 'объ_стр2_4_э': country_larg4_exp, 'объ_стр2_5_э':
               country_other1_exp, 'объ_стр2_6_э': itog_all_exp1, 'объ_стр1_1_э': country_larg1_1_exp, 'объ_стр1_2_э':
               country_larg1_2_exp, 'объ_стр1_3_э': country_larg1_3_exp, 'объ_стр1_4_э': country_larg1_4_exp,
           'объ_стр1_5_э': country_other2_exp, 'объ_стр1_6_э': itog_all_2_exp1, 'дин_стр1_э': dynamics_now_last1_exp,
           'дин_стр2_э': dynamics_now_last2_exp, 'дин_стр3_э': dynamics_now_last3_exp, 'дин_стр4_э':
               dynamics_now_last4_exp, 'дин_стр5_э': dynamics_now_last5_exp, 'умел_уменьш1_э': dynamics_text1_exp,
           'умел_уменьш2_э': dynamics_text2_exp, 'умел_уменьш3_э': dynamics_text3_exp, 'умел_уменьш4_э':
               dynamics_text4_exp, 'table_contents3_exp': table_contents3_exp, 'table_contents2_exp':
               table_contents2_exp, 'страна1_дол_э': country1_1_exp, 'страна2_дол_э': country2_1_exp, 'страна3_дол_э':
               country3_1_exp, 'страна4_дол_э': country4_1_exp, 'страна_доля_дол_э': country_part_2_exp,
           'table_contents_fo_exp': table_contents_fo_exp, 'фо1_э': fo1_exp, 'фо2_э': fo2_exp, 'фо3_э': fo3_exp,
           'фо4_э': fo4_exp, 'фо_доля_э': fo_sum_4_exp, 'table_contents_reg_exp': table_contents_reg_exp, 'рег1_э':
               reg1_exp, 'рег2_э': reg2_exp, 'рег3_э': reg3_exp, 'рег4_э': reg4_exp, 'рег_доля_э': reg_sum_4_exp,
           'table_fo_reg_exp': pivot_fo_reg_exp, 'объем_п': value_year_now, 'размерность_п': parameter_production,
           'значение_динамики_п': dynamics_now_last_prod, 'направление_динамики_п': line_dynamics_year_prod, 'объем3_п':
               value_production_m[year_int_2013], 'объем4_п': value_production_m[year_int_2014],
           'объем5_п': value_production_m[year_int_2015], 'объем6_п': value_production_m[year_int_2016], 'объем7_п':
               value_production_m[year_int_2017], 'объем8_п': value_production_m[year_int_2018], 'объем9_п':
               value_production_m[year_int_2019], 'объем10_п': value_production_m[year_int_2020], 'объем11_п':
               value_production_m[year_int_2021], 'image1_п': image_pro, 'рос1_п': dict_dynamics_year[year_int_2014],
           'рос2_п': dict_dynamics_year[year_int_2015], 'рос3_п': dict_dynamics_year[year_int_2016], 'рос4_п':
               dict_dynamics_year[year_int_2017], 'рос5_п': dict_dynamics_year[year_int_2018], 'рос6_п':
               dict_dynamics_year[year_int_2019], 'рос7_п': dict_dynamics_year[year_int_2020], 'рос8_п':
               dict_dynamics_year[year_int_2021], 'динамик_п': dynamics1_pro, 'процент1_п':
               dynamics_now_last_year_pro, 'год_макс_п': max_year_pro, 'объем_макс_п': value_max_year_pro, 'год_мин_п':
               min_year_pro, 'объем_мин_п': value_min_year_pro, 'table_contents_fo_pro':
               table_contents_fo_pro, 'фо_про': fo_1, 'фо_доля_про': fo_dol, 'table_contents_fo_din':
               table_contents_fo_din, 'объем_п_мод': value_production_now_year, 'размерность_пр_мод':
               parameter_production_now, 'объем_потр': consumption_dict[year_now], 'размерность_потр':
               dimension_consumption, 'об_пот3': consumption_dict[year2013], 'об_пот4': consumption_dict[year2014],
           'об_пот5': consumption_dict[year2015], 'об_пот6': consumption_dict[year2016], 'об_пот7':
               consumption_dict[year2017], 'об_пот8': consumption_dict[year2018], 'об_пот9': consumption_dict[year2019],
           'об_пот10': consumption_dict[year2020], 'об_пот11': consumption_dict[year2021], 'image4': image4, 'рос_пот1':
               variation_consumption[0][0], 'рос_пот2': variation_consumption[0][1], 'рос_пот3':
               variation_consumption[0][2], 'рос_пот4': variation_consumption[0][3], 'рос_пот5':
               variation_consumption[0][4], 'рос_пот6': variation_consumption[0][5], 'рос_пот7':
               variation_consumption[0][6], 'рос_пот8': variation_consumption[0][7], 'динамик_пот': dynamics4,
           'процент_потр': dynamics_netto_con, 'год_макс_пот': max_year_con, 'объем_макс_пот': value_max_year_con,
           'год_мин_пот': min_year_con, 'объем_мин_пот': value_min_year_con, 'об_и3': import_list_dict[year2013],
           'об_и4': import_list_dict[year2014], 'об_и5': import_list_dict[year2015], 'об_и6':
               import_list_dict[year2016], 'об_и7': import_list_dict[year2017], 'об_и8': import_list_dict[year2018],
           'об_и9': import_list_dict[year2019], 'об_и10': import_list_dict[year2020], 'об_и11':
               import_list_dict[year2021], 'об_э3': export_list_dict[year2013], 'об_э4': export_list_dict[year2014],
           'об_э5': export_list_dict[year2015], 'об_э6': export_list_dict[year2016], 'об_э7':
               export_list_dict[year2017], 'об_э8': export_list_dict[year2018], 'об_э9': export_list_dict[year2019],
           'об_э10': export_list_dict[year2020], 'об_э11': export_list_dict[year2021], 'об_п3':
               production_list_dict[year_int_2013], 'об_п4': production_list_dict[year_int_2014], 'об_п5':
               production_list_dict[year_int_2015], 'об_п6': production_list_dict[year_int_2016], 'об_п7':
               production_list_dict[year_int_2017], 'об_п8': production_list_dict[year_int_2018], 'об_п9':
               production_list_dict[year_int_2019], 'об_п10': production_list_dict[year_int_2020], 'об_п11':
               production_list_dict[year_int_2021], 'доля_про': part_pro, 'доля_им': part_im, 'image5': image5,
           'дол_им3': part_im_dict[year2013], 'дол_им4': part_im_dict[year2014], 'дол_им5': part_im_dict[year2015],
           'дол_им6': part_im_dict[year2016], 'дол_им7': part_im_dict[year2017], 'дол_им8': part_im_dict[year2018],
           'дол_им9': part_im_dict[year2019], 'дол_им10': part_im_dict[year2020], 'дол_им11': part_im_dict[year2021],
           'дин_дол_им': dinamic_im, 'размерность_п_м': parameter_production_m, 'ПРОДУКЦИЯ1': products1, 'сто_д3':
               dict_cost_dol[year2013], 'сто_д4': dict_cost_dol[year2014], 'сто_д5': dict_cost_dol[year2015], 'сто_д6':
               dict_cost_dol[year2016], 'сто_д7': dict_cost_dol[year2017], 'сто_д8': dict_cost_dol[year2018], 'сто_д9':
               dict_cost_dol[year2019], 'сто_д10': dict_cost_dol[year2020], 'сто_д11': dict_cost_dol[year2021],
           'дол_дин1': variation_cost_dol[0][0], 'дол_дин2': variation_cost_dol[0][1], 'дол_дин3':
               variation_cost_dol[0][2], 'дол_дин4': variation_cost_dol[0][3], 'дол_дин5': variation_cost_dol[0][4],
           'дол_дин6': variation_cost_dol[0][5], 'дол_дин7': variation_cost_dol[0][6], 'дол_дин8':
               variation_cost_dol[0][7], 'image6': image6, 'динамик_доллар': dynamics_cost, 'цифр_дин_доллар':
               dynamics_cost_val, 'мин_год_стоим_дол': min_year_cost, 'мин_стоим_дол': value_min_year_cost,
           'макс_год_стоим_дол': max_year_cost, 'макс_стоим_дол': value_max_year_cost, 'сто_руб3':
               ros_stoim_cost[year2013], 'сто_руб4': ros_stoim_cost[year2014], 'сто_руб5': ros_stoim_cost[year2015],
           'сто_руб6': ros_stoim_cost[year2016], 'сто_руб7': ros_stoim_cost[year2017], 'сто_руб8':
               ros_stoim_cost[year2018], 'сто_руб9': ros_stoim_cost[year2019], 'сто_руб10':
               ros_stoim_cost[year2020], 'сто_руб11': ros_stoim_cost[year2021], 'руб_дин1': variation_cost_rub[0][0],
           'руб_дин2': variation_cost_rub[0][1], 'руб_дин3': variation_cost_rub[0][2], 'руб_дин4': variation_cost_rub[0]
    [3], 'руб_дин5': variation_cost_rub[0][4], 'руб_дин6': variation_cost_rub[0][5], 'руб_дин7': variation_cost_rub[0]
    [6], 'руб_дин8': variation_cost_rub[0][7], 'image7': image7, 'динамик_рубль': dynamics_cost_rub, 'цифр_дин_рубль':
               dynamics_cost_val_rub, 'мин_год_стоим_руб': min_year_cost_rub, 'мин_стоим_руб': value_min_year_cost_rub,
        'макс_год_стоим_руб': max_year_cost_rub, 'макс_стоим_руб': value_max_year_cost_rub, 'table_country_dol':
               table_country_dol, 'table_country_dol_din': table_country_dol_din, 'table_country_rub':
               table_country_rub, 'сто_д_эк3': value_exp_cost[year2013], 'сто_д_эк4': value_exp_cost[year2014],
           'сто_д_эк5': value_exp_cost[year2015], 'сто_д_эк6': value_exp_cost[year2016], 'сто_д_эк7':
               value_exp_cost[year2017], 'сто_д_эк8': value_exp_cost[year2018], 'сто_д_эк9': value_exp_cost[year2019],
           'сто_д_эк10': value_exp_cost[year2020], 'сто_д_эк11': value_exp_cost[year2021], 'дол_дин_экс1':
               variation_cost_dol_exp[0][0], 'дол_дин_экс2': variation_cost_dol_exp[0][1], 'дол_дин_экс3':
               variation_cost_dol_exp[0][2], 'дол_дин_экс4': variation_cost_dol_exp[0][3], 'дол_дин_экс5':
               variation_cost_dol_exp[0][4], 'дол_дин_экс6': variation_cost_dol_exp[0][5], 'дол_дин_экс7':
               variation_cost_dol_exp[0][6], 'дол_дин_экс8': variation_cost_dol_exp[0][7], 'image8': image8,
           'динамик_доллар_экс': dynamics_cost_exp, 'цифр_дин_доллар_экс': dynamics_cost_val_exp,
           'мин_год_стоим_дол_экс': min_year_cost_exp, 'мин_стоим_дол_экс': value_min_year_cost_exp,
           'макс_год_стоим_дол_экс': max_year_cost_exp, 'макс_стоим_дол_экс': value_max_year_cost_exp, 'сто_руб_экс3':
               ros_stoim_cost_exp[year2013], 'сто_руб_экс4': ros_stoim_cost_exp[year2014], 'сто_руб_экс5':
               ros_stoim_cost_exp[year2015], 'сто_руб_экс6': ros_stoim_cost_exp[year2016], 'сто_руб_экс7':
               ros_stoim_cost_exp[year2017], 'сто_руб_экс8': ros_stoim_cost_exp[year2018], 'сто_руб_экс9':
               ros_stoim_cost_exp[year2019], 'сто_руб_экс10': ros_stoim_cost_exp[year2020], 'сто_руб_экс11':
               ros_stoim_cost_exp[year2021], 'руб_дин_экс1': variation_cost_rub_exp[0][0], 'руб_дин_экс2':
               variation_cost_rub_exp[0][1], 'руб_дин_экс3': variation_cost_rub_exp[0][2], 'руб_дин_экс4':
               variation_cost_rub_exp[0][3], 'руб_дин_экс5': variation_cost_rub_exp[0][4], 'руб_дин_экс6':
               variation_cost_rub_exp[0][5], 'руб_дин_экс7': variation_cost_rub_exp[0][6], 'руб_дин_экс8':
               variation_cost_rub_exp[0][7], 'image9': image9, 'динамик_рубль_экс': dynamics_cost_exp_rub,
           'цифр_дин_рубль_экс': dynamics_cost_val_exp_rub, 'мин_год_стоим_руб_экс': min_year_cost_exp_rub,
           'мин_стоим_руб_экс': value_min_year_cost_exp_rub, 'макс_год_стоим_руб_экс': max_year_cost_exp_rub,
           'макс_стоим_руб_экс': value_max_year_cost_exp_rub, 'table_country_dol_ex': table_country_dol_ex,
           'table_country_dol_din_ex': table_country_dol_din_ex, 'table_country_rub_ex': table_country_rub_ex,
           'код_вэд1': tnved1, 'код_вэд2': tnved2, 'код_вэд3': tnved3, 'наим_код_вэд1': name_tnved1, 'наим_код_вэд2':
               name_tnved2, 'наим_код_вэд3': name_tnved3}

doc.render(context)
doc.save(file_report_pattern_final)

wdFormatPDF = 17

in_file = os.path.abspath(file_report_pattern_final)
file_report_pdf = os.path.join('doc/report', f'{production}.pdf', )
out_file = os.path.abspath(file_report_pdf)

word = comtypes.client.CreateObject('Word.Application')
doc = word.Documents.Open(in_file)
doc.SaveAs(out_file, FileFormat=wdFormatPDF)
doc.Close()
word.Quit()

reader = PdfReader(out_file)
writer = PdfWriter()

writer.append_pages_from_reader(reader)
metadata = reader.metadata
writer.add_metadata(metadata)

# Write your custom metadata here:
writer.add_metadata({"/Author": "Бюро Готовых Исследований"})

with open(out_file, 'wb') as fp:
    writer.write(fp)
