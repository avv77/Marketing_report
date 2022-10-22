# настройки выгрузки таможни из базы данных

path_base = r'D:\База_Таможня\orders.db'
path_exel = r'D:\База_Таможня\Импорт.xlsx'
path_exel2 = r'D:\PyCharmProject\Marketing_report\exel\output2.xlsx'

list_year = ['2013', '2014', '2015', '2016', '2017', '2018', '2019', '2020', '2021']
tnved_1 = '3802100000'
# tnved_2 = '7004%'
# tnved_3 = '7005%'
# tnved_4 = '7006%'

list_name = ['NAPR', 'PERIOD', 'STRANA', 'TNVED', 'EDIZM', 'STOIM', 'NETTO', 'KOL', 'REGION', 'REGION_S']

# определяем размерность единиц (кг, тонны, тыс. тонн) 1- кг, 1000 - тонн, 1 000 000 тыс. тонн
dimension_netto = 1000000

# определяем размерность единиц ($, тыс. $, млн. $) 1- $, 1000 - тыс. $, 1 000 000  - млн. $
dimension_stoim = 1000000

# определяем размерность единиц (руб, тыс. руб., млн. руб., млрд. руб.) 1- руб., 1000 - тыс. руб., 1 000 000  -
# млн. руб., 1 000 000 000 - млрд. руб.
dimension_ros_stoim = 1000000

# год написания
year_now = '2021'
year_last = '2020'

# продукция
products1 = 'активированного угля'

# года написания
year2013 = '2013'
year2014 = '2014'
year2015 = '2015'
year2016 = '2016'
year2017 = '2017'
year2018 = '2018'
year2019 = '2019'
year2020 = '2020'
year2021 = '2021'

# курс доллара
rate = {
    '2013': 31.82,
    '2014': 38.33,
    '2015': 61.15,
    '2016': 66.96,
    '2017': 58.25,
    '2018': 62.78,
    '2019': 64.55,
    '2020': 72.13,
    '2021': 73.67
}

# файл для сохранения диаграмм в эксель
file_chart_exel = r'D:\PyCharmProject\Marketing_report\exel\bar.xlsx'

# файл для сохранения диаграммы как картинки
file_chart_png = r'D:\PyCharmProject\Marketing_report\png\chart.png'
file_chart_png2 = r'D:\PyCharmProject\Marketing_report\png\chart2.png'

# файл шаблон отчета
file_report_pattern = r'D:\PyCharmProject\Marketing_report\Импорт_пример.docx'

