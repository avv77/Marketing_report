# настройки выгрузки таможни из базы данных

path_base = r'D:\База_Таможня\orders.db'
path_exel = r'D:\База_Таможня\Импорт.xlsx'
path_exel2 = r'D:\PyCharmProject\Marketing_report\exel\output2.xlsx'
path_produst = r'D:\PyCharmProject\Marketing_report\exel\1_4_production.xlsx'

list_year = ['2013', '2014', '2015', '2016', '2017', '2018', '2019', '2020', '2021']

# три кода тн вэд максимум, если больше - внести изменение в функцию import_from_base (custom_selection.py)

list_name = ['NAPR', 'PERIOD', 'STRANA', 'TNVED', 'EDIZM', 'STOIM', 'NETTO', 'KOL', 'REGION', 'REGION_S']

# определяем размерность единиц (кг, тонны, тыс. тонн) 1- кг, 1000 - тонн, 1000000 - тыс. тонн
dimension_netto = 1000000

# определяем размерность единиц ($, тыс. $, млн. $) 1- $, 1000 - тыс. $, 1000000  - млн. $
dimension_stoim = 1000000

# определяем размерность единиц (руб, тыс. руб., млн. руб., млрд. руб.) 1- руб., 1000 - тыс. руб., 1000000  -
# млн. руб., 1000000000 - млрд. руб. Должна быть равна предыдущему значению.
dimension_ros_stoim = 1000000

# год написания
year_now = '2021'
year_last = '2020'

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

# года написания
year_int_2013 = 2013
year_int_2014 = 2014
year_int_2015 = 2015
year_int_2016 = 2016
year_int_2017 = 2017
year_int_2018 = 2018
year_int_2019 = 2019
year_int_2020 = 2020
year_int_2021 = 2021

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
file_chart_png3 = r'D:\PyCharmProject\Marketing_report\png\chart3.png'
file_chart_png4 = r'D:\PyCharmProject\Marketing_report\png\chart4.png'
file_chart_png5 = r'D:\PyCharmProject\Marketing_report\png\chart5.png'
file_chart_png6 = r'D:\PyCharmProject\Marketing_report\png\chart6.png'
file_chart_png7 = r'D:\PyCharmProject\Marketing_report\png\chart7.png'
file_chart_png8 = r'D:\PyCharmProject\Marketing_report\png\chart8.png'
file_chart_png9 = r'D:\PyCharmProject\Marketing_report\png\chart9.png'
file_chart_png10 = r'D:\PyCharmProject\Marketing_report\png\chart10.png'
file_chart_png11 = r'D:\PyCharmProject\Marketing_report\png\chart10=1.png'

# файл шаблон отчета
file_report_pattern = r'D:\PyCharmProject\Marketing_report\doc\report\Импорт_пример.docx'

# файл заполненный шаблон отчета
file_report_pattern_final = r'doc\report\Отчет маркетингового исследования.docx'

# файл заполненный шаблон отчета в pdf
file_report_pdf = r'doc\report'

# название продукции из файла по производству
production = 'Антидетонаторы; присадки к топливу и смазочным материалам и аналогичные продукты'

# продукция
products1 = 'антидетонаторов'

# индекс для приведение к размерности производства, импорта и экспорта
index_import = 1
index_export = 1
index_production = 1000

# при ошибке module 'win32com.gen_py.00020813-0000-0000-C000-000000000046x0x1x9' has no attribute 'CLSIDToClassMap'
# удалить каталог по адресу C:\Users\mag77\AppData\Local\Temp\gen_py
