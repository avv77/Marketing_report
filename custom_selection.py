import sqlite3
import openpyxl
import pandas as pd
from file_import import run_import_exel, size_netto, size_stoim, size_ros_stoim
from settings import path_exel, list_year, tnved_1, list_name, path_base, year_now, year_last, dimension_netto, \
    dimension_stoim, dimension_ros_stoim


def import_from_base(path, exp_im):
    conn = sqlite3.connect(path)
    writer = pd.ExcelWriter(path_exel)
    for year in list_year:
        cur = conn.cursor()
        cur.execute(f'SELECT * FROM "{year}" WHERE NAPR = "{exp_im}" AND TNVED LIKE "{tnved_1}"')
        result = cur.fetchall()
        daf = pd.DataFrame(result)
        daf.to_excel(writer, year, index=False)
        writer.save()
        print(f'Обработана база {year}')
    wb = openpyxl.load_workbook(path_exel)
    for year in list_year:
        number = 1
        sheet = wb[year]
        for name in list_name:
            cell = sheet.cell(row=1, column=number)
            cell.value = name
            number += 1
    wb.save(path_exel)
    print('Выгрузка закончена')


def import_base_processing(im_exp):
    # выгружаем импорт по коду из базы данных таможни
    import_from_base(path_base, im_exp)

    # составляем таблицу из веса и стоимости по годам в виде словаря данных

    dict_import = run_import_exel()
    dict_netto = dict_import[0]
    dict_stoim = dict_import[1]

    # определяем показатели для записи первого абзаца
    dict_netto_year_now = dict_netto[year_now]
    dict_netto_year_last = dict_netto[year_last]

    dynamics = round(dict_netto_year_now / dict_netto_year_last * 100 - 100, 1)
    if dynamics >= 0:
        direction = 'больше'
    else:
        direction = 'меньше'

    dict_stoim_year_now = dict_stoim[year_now]
    dict_stoim_year_last = dict_stoim[year_last]

    # определяем размерность показателей
    size_netto_now = size_netto(dimension_netto)
    size_stoim_now = size_stoim(dimension_stoim)
    size_ros_stoim_now = size_ros_stoim(dimension_ros_stoim)
    return dict_netto_year_now, dynamics, direction, size_netto_now, dict_netto, dict_stoim_year_now, \
           dict_stoim_year_last, size_stoim_now, dict_stoim, size_ros_stoim_now


def variation(dict_netto):
    list_variation_probe = []
    list_variation = []
    for key in dict_netto.keys():
        list_variation_probe.append(dict_netto[key])
    len_list_variation_probe = len(list_variation_probe)
    for netto_number in range(len_list_variation_probe):
        if netto_number < len_list_variation_probe - 1:
            variation_value = list_variation_probe[netto_number + 1] / list_variation_probe[netto_number] * 100 - 100
            list_variation.append(round(variation_value, 1))
    return list_variation, list_variation_probe


def variation_stoim(dict_stoim):
    list_variation_probe = []
    list_variation = []
    for key in dict_stoim.keys():
        list_variation_probe.append(dict_stoim[key])
    len_list_variation_probe = len(list_variation_probe)
    for netto_number in range(len_list_variation_probe):
        if netto_number < len_list_variation_probe - 1:
            variation_value = list_variation_probe[netto_number + 1] / list_variation_probe[netto_number] * 100 - 100
            list_variation.append(round(variation_value, 1))
    return list_variation, list_variation_probe


def variation_stoim_ros(dict_stoim_ros):
    list_variation_probe = []
    list_variation = []
    for key in dict_stoim_ros.keys():
        list_variation_probe.append(dict_stoim_ros[key])
    len_list_variation_probe = len(list_variation_probe)
    for netto_number in range(len_list_variation_probe):
        if netto_number < len_list_variation_probe - 1:
            variation_value = list_variation_probe[netto_number + 1] / list_variation_probe[netto_number] * 100 - 100
            list_variation.append(round(variation_value, 1))
    return list_variation, list_variation_probe