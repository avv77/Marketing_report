import decimal
import pandas
from pivot_table_country import moneyfmt
from settings import dimension_stoim, path_exel, rate, dimension_netto


def sum_column_exel(list1, dimension):
    total = 0
    for cell in list1:
        if type(cell) == str:
            cell = cell.replace(",", ".")
            cell = float(cell)
            cell = cell / dimension
            total += cell
        else:
            cell = cell / dimension
            total += cell
    return total


def all_sum_dict(sheet, table_stoim_netto, stoim_netto, dimension_stoim_netto):
    file_import2 = pandas.read_excel(path_exel, sheet)
    stoim_list = file_import2[stoim_netto].tolist()
    total1 = sum_column_exel(stoim_list, dimension_stoim_netto)
    if dimension_stoim == 1000000:
        table_stoim_netto[sheet] = round(total1, 1)
    else:
        table_stoim_netto[sheet] = round(total1)


def size_netto(dimension_netto):
    size_netto1 = ''
    if dimension_netto == 1:
        size_netto1 = 'кг'
    elif dimension_netto == 1000:
        size_netto1 = 'тонн'
    elif dimension_netto == 1000000:
        size_netto1 = 'тыс. тонн'
    return size_netto1


def size_stoim(dimension_stoim1):
    size_stoim1 = ''
    if dimension_stoim1 == 1:
        size_stoim1 = '$'
    elif dimension_stoim1 == 1000:
        size_stoim1 = 'тыс. $'
    elif dimension_stoim1 == 1000000:
        size_stoim1 = 'млн. $'
    return size_stoim1


def size_ros_stoim(dimension_ros_stoim):
    size_stoim2 = ''
    if dimension_ros_stoim == 1:
        size_stoim2 = 'руб'
    elif dimension_ros_stoim == 1000:
        size_stoim2 = 'тыс. руб'
    elif dimension_ros_stoim == 1000000:
        size_stoim2 = 'млн. руб'
    elif dimension_ros_stoim == 1000000000:
        size_stoim2 = 'млрд. руб'
    return size_stoim2


def run_import_exel():
    file_import = pandas.ExcelFile(path_exel)
    sheets = file_import.sheet_names
    table_stoim = {}
    table_netto = {}

    for sheet in sheets:
        all_sum_dict(sheet, table_stoim, 'STOIM', dimension_stoim)
        all_sum_dict(sheet, table_netto, 'NETTO', dimension_netto)
    return [table_netto, table_stoim]


def dynamics(a, b):
    if b > a:
        dynamics1 = 'выросли'
    elif b < a:
        dynamics1 = 'сократились'
    else:
        dynamics1 = 'не изменились'
    return dynamics1


def ros_stoim_table(dict_dollar, dimension):
    dict_ros_stoim = {}
    for key in dict_dollar.keys():
        if dimension_stoim == 1:
            if dimension == 1:
                dict_ros_stoim[key] = round(dict_dollar[key] * rate[key])
            elif dimension == 1000:
                dict_ros_stoim[key] = round(dict_dollar[key] * rate[key] / 1000)
            elif dimension == 1000000:
                dict_ros_stoim[key] = round(dict_dollar[key] * rate[key] / 1000000)
            elif dimension == 1000000000:
                dict_ros_stoim[key] = round(dict_dollar[key] * rate[key] / 1000000000)
        if dimension_stoim == 1000:
            if dimension == 1:
                dict_ros_stoim[key] = round(dict_dollar[key] * rate[key] * 1000)
            elif dimension == 1000:
                dict_ros_stoim[key] = round(dict_dollar[key] * rate[key])
            elif dimension == 1000000:
                dict_ros_stoim[key] = round(dict_dollar[key] * rate[key] / 1000)
            elif dimension == 1000000000:
                dict_ros_stoim[key] = round(dict_dollar[key] * rate[key] / 1000000)
        if dimension_stoim == 1000000:
            if dimension == 1:
                dict_ros_stoim[key] = round(dict_dollar[key] * rate[key] * 1000000)
            elif dimension == 1000:
                dict_ros_stoim[key] = round(dict_dollar[key] * rate[key] * 1000)
            elif dimension == 1000000:
                dict_ros_stoim[key] = round(dict_dollar[key] * rate[key])
            elif dimension == 1000000000:
                dict_ros_stoim[key] = round(dict_dollar[key] * rate[key] / 1000)
    return dict_ros_stoim


def ros_stoim_table_cost(dict_dollar):
    dict_ros_stoim = {}
    for key in dict_dollar.keys():
            dict_ros_stoim[key] = round(dict_dollar[key] * rate[key])
    return dict_ros_stoim


def transformation_value(value):
    d = decimal.Decimal(value)
    d = moneyfmt(d, sep=' ')
    return d


def dynamics_value_year(dynamics_value):
    if dynamics_value >= 0:
        direction = 'больше'
    else:
        direction = 'меньше'
    return direction


def cost_dol():
    file_import = pandas.ExcelFile(path_exel)
    sheets = file_import.sheet_names
    table_stoim = {}
    table_netto = {}

    for sheet in sheets:
        file_import2 = pandas.read_excel(path_exel, sheet)
        stoim_list = file_import2['STOIM'].tolist()
        total1 = sum_column_exel(stoim_list, 1)
        table_stoim[sheet] = round(total1)

        file_import2 = pandas.read_excel(path_exel, sheet)
        stoim_list = file_import2['NETTO'].tolist()
        total2 = sum_column_exel(stoim_list, 1)
        table_netto[sheet] = round(total2)

    netto = table_netto
    stoim = table_stoim

    cost = {}
    for year, number in netto.items():
        if netto[year] == 0:
            cost[year] = 0
        else:
            cost[year] = round(stoim[year] / netto[year], 2)
    return cost
