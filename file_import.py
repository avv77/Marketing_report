import decimal
import pandas
from pivot_table_country import moneyfmt
from settings import dimension_stoim, path_exel, rate


def sum_column_exel(list, dimension):
    total = 0
    for cell in list:
        if type(cell) == str:
            cell = cell.replace(",", ".")
            cell = float(cell)
            cell = cell / dimension
            total += cell
        else:
            cell = cell / dimension
            total += cell
    return total


def all_sum_dict(sheet, table_stoim_netto, stoim_netto):
    file_import2 = pandas.read_excel(path_exel, sheet)
    stoim_list = file_import2[stoim_netto].tolist()
    total1 = sum_column_exel(stoim_list, dimension_stoim)
    if dimension_stoim == 1000000:
        table_stoim_netto[sheet] = round(total1, 1)
    else:
        table_stoim_netto[sheet] = round(total1)


def size_netto(dimension_netto):
    size_netto = ''
    if dimension_netto == 1:
        size_netto = 'кг'
    elif dimension_netto == 1000:
        size_netto = 'тонн'
    elif dimension_netto == 1000000:
        size_netto = 'тыс. тонн'
    return size_netto


def size_stoim(dimension_stoim):
    size_stoim = ''
    if dimension_stoim == 1:
        size_stoim = '$'
    elif dimension_stoim == 1000:
        size_stoim = 'тыс. $'
    elif dimension_stoim == 1000000:
        size_stoim = 'млн. $'
    return size_stoim


def size_ros_stoim(dimension_ros_stoim):
    size_stoim = ''
    if dimension_ros_stoim == 1:
        size_stoim = 'руб'
    elif dimension_ros_stoim == 1000:
        size_stoim = 'тыс. руб'
    elif dimension_ros_stoim == 1000000:
        size_stoim = 'млн. руб'
    elif dimension_ros_stoim == 1000000000:
        size_stoim = 'млрд. руб'
    return size_stoim


def run_import_exel():
    file_import = pandas.ExcelFile(path_exel)
    sheets = file_import.sheet_names
    table_stoim = {}
    table_netto = {}

    for sheet in sheets:
        all_sum_dict(sheet, table_stoim, 'STOIM')
        all_sum_dict(sheet, table_netto, 'NETTO')
    return [table_netto, table_stoim]


def dynamics(a, b):
    if b > a:
        dynamics1 = 'выросли'
    elif b < a:
        dynamics1 = 'сократились'
    else:
        dynamics1 = 'не изменились'
    return dynamics1


def ros_stoim_table(dict_dollar):
    dict_ros_stoim = {}
    for key in dict_dollar.keys():
        dict_ros_stoim[key] = round(dict_dollar[key] * rate[key])
    return dict_ros_stoim


def transformation_value(value):
    d = decimal.Decimal(value)
    d = moneyfmt(d, sep=' ')
    return d
