import decimal
from docxtpl import DocxTemplate
import pandas as pd
from openpyxl import load_workbook

from pivot_table_country import moneyfmt
from settings import path_exel2, file_report_pattern


def pivot_table_fo_reg(path_fo_reg, pivot_index, pivot_values, pivot_1_column, pivot_2_column, path_exel_last):
    data_frame_country = pd.read_excel(path_exel_last)
    report_table = data_frame_country.pivot_table(index=pivot_index, values=pivot_values, aggfunc='sum', margins=True).\
        round(0)
    report_table.to_excel(path_fo_reg, sheet_name='Report')

    table_contents_country = []
    workbook = load_workbook(path_fo_reg)
    sheet_1 = workbook['Report']
    values_fo_last = ''
    for i in range(2, sheet_1.max_row + 1):
        table_year_dict = dict()
        values_fo = sheet_1.cell(i, 1).value
        if values_fo is None:
            table_year_dict[pivot_1_column] = values_fo_last
        else:
            table_year_dict[pivot_1_column] = values_fo
            values_fo_last = values_fo
        table_year_dict[pivot_2_column] = sheet_1.cell(i, 2).value
        for j in range(3, sheet_1.max_column + 1):
            d = sheet_1.cell(i, j).value
            if d is not None:
                d = decimal.Decimal(sheet_1.cell(i, j).value)
                d = moneyfmt(d, sep=' ')
                table_year_dict[sheet_1.cell(1, j).value] = d
        table_contents_country.append(table_year_dict)

    return table_contents_country


i = pivot_table_fo_reg(r'exel\report_fo_reg.xlsx', ['ФО', 'Регионы'], 'NETTO', 'ФО', 'Регионы', path_exel2)
context = {'table_fo_reg': i}
doc = DocxTemplate(file_report_pattern)
doc.render(context)
doc.save(r'exel\Импорт_пример_финал.docx')
