import win32com.client
from PIL import ImageGrab
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from settings import file_chart_exel, file_chart_png, file_chart_png2
import matplotlib.pyplot as plt


def chart1(dict_netto, size_netto_now):

    wb = Workbook()
    ws = wb.active

    # список столбцов и строк
    dict_list = [[key, value] for key, value in dict_netto.items()]
    rows_chart = [['год', 'объем']]

    for i in dict_list:
        rows_chart.append(i)

    for row in rows_chart:
        ws.append(row)

    # создаем объект диаграммы
    chart1 = BarChart()
    # установим тип - `вертикальные столбцы`
    chart1.type = "col"
    # установим стиль диаграммы (цветовая схема)
    chart1.style = 10
    # заголовок диаграммы
    chart1.title = ""
    # подпись оси `y`
    chart1.y_axis.title = size_netto_now
    # показывать данные на оси (для LibreOffice Calc)
    chart1.y_axis.delete = False
    # подпись оси `x`
    chart1.x_axis.title = 'год'
    chart1.x_axis.delete = False
    # отключим линии сетки
    chart1.y_axis.majorGridlines = None
    # уберем легенду
    chart1.legend = None
    # выберем 2 столбца с данными для оси `y`
    data = Reference(ws, min_col=2, max_col=2, min_row=1, max_row=10)
    # теперь выберем категорию для оси `x`
    categor = Reference(ws, min_col=1, min_row=2, max_row=10)
    # добавляем данные в объект диаграммы
    chart1.add_data(data, titles_from_data=True)
    # установим метки на объект диаграммы
    chart1.set_categories(categor)

    s1 = chart1.series[0]
    s1.marker.symbol = "diamond"
    chart1.dataLabels = DataLabelList()
    chart1.dataLabels.showVal = True

    # добавим диаграмму на лист, в ячейку "D10"
    ws.add_chart(chart1, "D10")
    wb.save(file_chart_exel)

    input_file = file_chart_exel
    output_image = file_chart_png

    operation = win32com.client.Dispatch("Excel.Application")
    operation.Visible = 0
    operation.DisplayAlerts = 0
    workbook_2 = operation.Workbooks.Open(input_file)
    sheet_2 = operation.Sheets(1)

    for x, chart in enumerate(sheet_2.Shapes):
        chart.Copy()
        image = ImageGrab.grabclipboard()
        image.save(output_image, 'png')
        pass
    workbook_2.Close(True)
    operation.Quit()
    pass


# круговая диаграмма
def chart2(table_contents_country):

    # список столбцов и строк
    country_name_list = []
    country_value_list = []

    itog = 0
    for i in range(len(table_contents_country) - 2):
        itog += float(table_contents_country[i]['NETTO'])

    other = round(itog - table_contents_country[0]['NETTO'] - table_contents_country[1]['NETTO'] -
                  table_contents_country[2]['NETTO'] - table_contents_country[3]['NETTO'])

    for i_value in range(4):
        country_value_list.append(table_contents_country[i_value]['NETTO'])
    country_value_list.append(other)

    for i_name in range(4):
        country_name_list.append(table_contents_country[i_name]['Страны'])
    country_name_list.append('другие')

    fig1, ax1 = plt.subplots()
    explode = (0.1, 0, 0, 0, 0)
    ax1.pie(country_value_list, explode=explode, labels=country_name_list, autopct='%1.1f%%', shadow=True,
            startangle=90)
    fig1.savefig(file_chart_png2)
    pass
