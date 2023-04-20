import win32com.client
from PIL import ImageGrab
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from settings import file_chart_exel
import matplotlib.pyplot as plt


def chart1(dict_netto, size_netto_now, file_png):

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
    chart1_1 = BarChart()
    # установим тип - `вертикальные столбцы`
    chart1_1.type = "col"
    # установим стиль диаграммы (цветовая схема)
    chart1_1.style = 10
    # заголовок диаграммы
    chart1_1.title = ""
    # подпись оси `y`
    chart1_1.y_axis.title = size_netto_now
    # показывать данные на оси (для LibreOffice Calc)
    chart1_1.y_axis.delete = False
    # подпись оси `x`
    chart1_1.x_axis.title = 'год'
    chart1_1.x_axis.delete = False
    # отключим линии сетки
    chart1_1.y_axis.majorGridlines = None
    # уберем легенду
    chart1_1.legend = None
    # выберем 2 столбца с данными для оси `y`
    data = Reference(ws, min_col=2, max_col=2, min_row=1, max_row=10)
    # теперь выберем категорию для оси `x`
    categor = Reference(ws, min_col=1, min_row=2, max_row=10)
    # добавляем данные в объект диаграммы
    chart1_1.add_data(data, titles_from_data=True)
    # установим метки на объект диаграммы
    chart1_1.set_categories(categor)

    s1 = chart1_1.series[0]
    s1.marker.symbol = "diamond"
    chart1_1.dataLabels = DataLabelList()
    chart1_1.dataLabels.showVal = True

    # добавим диаграмму на лист, в ячейку "D10"
    ws.add_chart(chart1_1, "D10")
    wb.save(file_chart_exel)

    input_file = file_chart_exel
    output_image = file_png

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


# круговая диаграмма по 4-м крупнейшим
def chart2(table_contents_country, file_png):

    # список столбцов и строк
    country_name_list = []
    country_value_list = []

    if table_contents_country[0]['Доля'] < 80:
        itog = 0
        for i in range(len(table_contents_country) - 1):
            itog += float(table_contents_country[i]['NETTO'])

        if len(table_contents_country) < 6:
            for i_value in range(len(table_contents_country) - 1):
                country_value_list.append(table_contents_country[i_value]['NETTO'])
            for i_name in range(len(table_contents_country) - 1):
                country_name_list.append(table_contents_country[i_name]['Страны'])
        else:
            other = round(itog - table_contents_country[0]['NETTO'] - table_contents_country[1]['NETTO'] -
                          table_contents_country[2]['NETTO'] - table_contents_country[3]['NETTO'])
            for i_value in range(4):
                country_value_list.append(table_contents_country[i_value]['NETTO'])
            country_value_list.append(other)

            for i_name in range(4):
                country_name_list.append(table_contents_country[i_name]['Страны'])
            country_name_list.append('другие')

        total = sum(country_value_list)
        labels = [f"{n} ({v / total:.1%})" for n, v in zip(country_name_list, country_value_list)]
        fig1, ax1 = plt.subplots()
        ax1.pie(country_value_list, radius=1.1, explode=[0.15] + [0 for _ in range(len(country_name_list) - 1)])
        ax1.legend(bbox_to_anchor=(-0.3, 0), loc=3, labels=labels, borderaxespad=0, ncol=3)
        fig1.savefig(file_png)

    else:
        itog = 0
        if len(table_contents_country) == 2:
            itog = float(table_contents_country[0]['NETTO'])
        elif len(table_contents_country) == 3:
            itog = float(table_contents_country[0]['NETTO']) + float(table_contents_country[1]['NETTO'])
        elif len(table_contents_country) == 4:
            itog = float(table_contents_country[0]['NETTO']) + float(table_contents_country[1]['NETTO']) + float(
                table_contents_country[2]['NETTO'])
        elif len(table_contents_country) == 5:
            itog = float(table_contents_country[0]['NETTO']) + float(table_contents_country[1]['NETTO']) + float(
                table_contents_country[2]['NETTO']) + float(table_contents_country[3]['NETTO'])
        else:
            for i in range(len(table_contents_country) - 2):
                itog += float(table_contents_country[i]['NETTO'])

        other = round(itog - table_contents_country[0]['NETTO'])

        country_value_list.append(table_contents_country[0]['NETTO'])
        country_value_list.append(other)

        country_name_list.append(table_contents_country[0]['Страны'])
        country_name_list.append('другие')

        total = sum(country_value_list)
        labels = [f"{n} ({v / total:.1%})" for n, v in zip(country_name_list, country_value_list)]
        fig1, ax1 = plt.subplots()
        ax1.pie(country_value_list, radius=1.1, explode=[0.15] + [0 for _ in range(len(country_name_list) - 1)])
        ax1.legend(bbox_to_anchor=(-0.3, 0), loc=3, labels=labels, borderaxespad=0, ncol=3)
        fig1.savefig(file_png)
    pass


# круговая диаграмма (table_name - список наименований, table_value - список значений)

def chart3(table_name, table_value, file_png):

    fig1, ax1 = plt.subplots()
    ax1.pie(table_value, explode=None, labels=table_name, autopct='%1.1f%%', shadow=True, startangle=90)
    fig1.savefig(file_png)
    pass
