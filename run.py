from docxtpl import DocxTemplate, InlineImage
from chart1 import chart1, chart2
from custom_selection import import_base_processing, variation, variation_stoim, variation_stoim_ros
from file_import import dynamics, ros_stoim_table, transformation_value
from pivot_table_country import pivot1_1, pivot1_2, pivot1_3, pivot2_1, pivot2_2, pivot_table_country_year, pivotfo, \
    pivot_fo_1, pivot_table_fo_reg
from settings import year_now, year_last, products1, year2013, year2014, year2015, year2016, year2017, year2018, \
    year2019, year2020, year2021, file_report_pattern, file_chart_png2, list_year, path_exel2
from docx.shared import Cm
from settings import file_chart_png
import pandas as pd

import_base = import_base_processing('ИМ')
doc = DocxTemplate(file_report_pattern)
chart1(import_base[4], import_base[3])
image = InlineImage(doc, file_chart_png, Cm(16.5))
variation_value = variation(import_base[4])
dynamics_netto = round(variation_value[1][len(variation_value[1]) - 1] / variation_value[1][0] * 100 - 100, 1)
max_year = max(import_base[4], key=import_base[4].get)
value_max_year = import_base[4][max_year]
min_year = min(import_base[4], key=import_base[4].get)
value_min_year = import_base[4][min_year]
variation_value_stoim = variation_stoim(import_base[8])
dynamics_stoim = round(variation_value_stoim[1][len(variation_value_stoim[1]) - 1] /
                       variation_value_stoim[1][0] * 100 - 100, 1)
dynamics1 = dynamics(variation_value[1][0], variation_value[1][len(variation_value[1]) - 1])
max_year_stoim = max(import_base[8], key=import_base[8].get)
value_max_year_stoim = import_base[8][max_year_stoim]
min_year_stoim = min(import_base[8], key=import_base[8].get)
value_min_year_stoim = import_base[8][min_year_stoim]
dynamics2 = dynamics(variation_value_stoim[1][0], variation_value_stoim[1][len(variation_value_stoim[1]) - 1])
# заполняем таблицу - стоимость в рублях
ros_stoim = ros_stoim_table(import_base[8])
# заполняем таблицу - динамика стоимости в рублях
variation_value_stoim_ros = variation_stoim_ros(ros_stoim)
# заполняем текстовку после таблицы
dynamics3 = dynamics(variation_value_stoim_ros[1][0], variation_value_stoim_ros[1][len(variation_value_stoim_ros[1]) -
                                                                                   1])
dynamics_stoim_ros = round(variation_value_stoim_ros[1][len(variation_value_stoim[1]) - 1] /
                       variation_value_stoim_ros[1][0] * 100 - 100, 1)
max_year_stoim_ros = max(ros_stoim, key=ros_stoim.get)
value_max_year_stoim_ros = ros_stoim[max_year_stoim_ros]
min_year_stoim_ros = min(ros_stoim, key=ros_stoim.get)
value_min_year_stoim_ros = ros_stoim[min_year_stoim_ros]

# заполняем сводную таблицу по странам в кг
pivot1_1(list_year[len(list_year) - 1])
pivot1_2()
pivot3 = pivot1_3()
table_contents = pivot3[0]
country1 = table_contents[0]['Страны']
country2 = table_contents[1]['Страны']
country3 = table_contents[2]['Страны']
country4 = table_contents[3]['Страны']
country_part = pivot3[1]

# делаем и вставляем круговую диаграмму по странам
chart2(pivot3[2])
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

country_other1 = round(itog_all_ - pivot3[2][0]['NETTO'] - pivot3[2][1]['NETTO'] -
              pivot3[2][2]['NETTO'] - pivot3[2][3]['NETTO'])
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

country_other2 = round(itog_all_2_ - pivot3_2[2][0]['NETTO'] - pivot3_2[2][1]['NETTO'] -
              pivot3_2[2][2]['NETTO'] - pivot3_2[2][3]['NETTO'])
country_other2 = transformation_value(country_other2)

dynamics_now_last1 = round(country_larg1_tr / country_larg1_1_tr * 100 - 100, 1)
dynamics_now_last2 = round(country_larg2_tr / country_larg1_2_tr * 100 - 100, 1)
dynamics_now_last3 = round(country_larg3_tr / country_larg1_3_tr * 100 - 100, 1)
dynamics_now_last4 = round(country_larg4_tr / country_larg1_4_tr * 100 - 100, 1)
dynamics_now_last5 = round(itog_all_ / itog_all_2_ * 100 - 100, 1)

dynamics_text1 = dynamics(country_larg1_1_tr, country_larg1_tr)
dynamics_text2 = dynamics(country_larg1_2_tr, country_larg2_tr)
dynamics_text3 = dynamics(country_larg1_3_tr, country_larg3_tr)
dynamics_text4 = dynamics(country_larg1_4_tr, country_larg4_tr)

year_contrast1 = list_year[len(list_year) - 2]
year_contrast2 = list_year[len(list_year) - 1]

# заполняем таблицу со странами по стоимости
pivot1_1(list_year[len(list_year) - 1])
pivot2_1()
pivot4 = pivot2_2()
table_contents2 = pivot4[0]
country1_1 = table_contents2[0]['Страны']
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
fo2 = table_contents_fo[1]['ФО']
fo3 = table_contents_fo[2]['ФО']
fo4 = table_contents_fo[3]['ФО']
fo_sum_4 = pivot_fo[1]

# заполняем таблицу с регионами в кг
pivot_fo_1('Регионы', 'NETTO', r'exel\report_reg.xlsx')
pivot_reg = pivotfo(r'exel\report_reg.xlsx', 'Регионы', 'NETTO')
table_contents_reg = pivot_reg[0]
reg1 = table_contents_reg[0]['Регионы']
reg2 = table_contents_reg[1]['Регионы']
reg3 = table_contents_reg[2]['Регионы']
reg4 = table_contents_reg[3]['Регионы']
reg_sum_4 = pivot_reg[1]

# заполняем таблицу по федеральным округам и регионам (4.12)
pivot_fo_reg = pivot_table_fo_reg(r'exel\report_fo_reg.xlsx', ['ФО', 'Регионы'], 'NETTO', 'ФО', 'Регионы', path_exel2)

# заполняем динамику по экспорту
export_base = import_base_processing('ЭК')


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
               dynamics2, 'размерность3': import_base[9], 'руб3': ros_stoim['2013'], 'руб4': ros_stoim['2014'], 'руб5':
               ros_stoim['2015'], 'руб6': ros_stoim['2016'], 'руб7': ros_stoim['2017'], 'руб8': ros_stoim['2018'],
           'руб9': ros_stoim['2019'], 'руб10': ros_stoim['2020'], 'руб11': ros_stoim['2021'], 'бол1':
               variation_value_stoim_ros[0][0], 'бол2': variation_value_stoim_ros[0][1], 'бол3':
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
           'itog_all_2': country_other2, 'дин_стр1': dynamics_now_last1, 'дин_стр2': dynamics_now_last2, 'дин_стр3':
               dynamics_now_last3, 'дин_стр4': dynamics_now_last4, 'дин_стр5': dynamics_now_last5, 'умел_уменьш1':
               dynamics_text1, 'умел_уменьш2': dynamics_text2, 'умел_уменьш3': dynamics_text3, 'умел_уменьш4':
               dynamics_text4, 'table_contents3': table_contents3, 'table_contents_fo': table_contents_fo, 'фо1': fo1,
           'фо2': fo2, 'фо3': fo3, 'фо4': fo4, 'фо_доля': fo_sum_4, 'table_contents_reg': table_contents_reg, 'рег1':
               reg1, 'рег2': reg2, 'рег3': reg3, 'рег4': reg4, 'рег_доля': reg_sum_4,
           'table_fo_reg': pivot_fo_reg, 'объем_э': export_base[0], 'размерность_э': export_base[3],
           'значение_динамики_э': export_base[1], 'направление_динамики_э': export_base[2],
           'экс3': export_base[4]['2013'], 'экс4': export_base[4]['2014'], 'экс5': export_base[4]['2015'], 'экс6':
               export_base[4]['2016'], 'экс7': export_base[4]['2017'], 'экс8': export_base[4]['2018'], 'экс9':
               export_base[4]['2019'], 'экс10': export_base[4]['2020'], 'экс11': export_base[4]['2021']}
doc.render(context)
doc.save(r'exel\Импорт_пример_финал.docx')
