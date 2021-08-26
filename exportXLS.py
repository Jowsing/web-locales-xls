# -*- coding: UTF-8 -*-

import sys
import os
import xlwt
import shutil

root_path = sys.path[0]


def r_newline(str):
    return str.replace('\n', '').replace('\r', '')


def z_split(str, sls, srs):
    sls_len = len(sls)
    if sls_len != len(srs):
        return
    for i in range(sls_len):
        (sl, sr) = (sls[i], srs[i])
        str = str.replace(sl + sr, sl+'<->')
    return str.split('<->')


def z_arr_push(arr, value):
    i = 0
    if value in arr:
        i = arr.index(value)
    else:
        arr.append(value)
        i = len(arr) - 1
    return (i, arr)


def filter_dir(file):
    return not file.__contains__('.')


def filter_js_file(file):
    return file.__contains__('.js')


def copy_dir(from_path, to_path):
    shutil.copytree(from_path, to_path)


def search_files(path, filter):
    files = os.listdir(path)
    result = []
    for file in files:
        if filter(file):
            result.append(file)
    return result


def read_js_json(js_file):
    with open(js_file, 'r') as file:
        js_str = file.read()
        js_str = js_str[js_str.find('{') + 1: js_str.find('}')]  # 去除 {}
        js_str = r_newline(js_str)  # 去除换行
        local_str_list = z_split(js_str, ['\'', '`'], [',', ','])
        local_kv_list = []
        for local_str in local_str_list:
            if len(local_str) == 0:
                continue
            str = r_newline(local_str.strip())
            [key, value] = z_split(str, '\'', ':')
            if key.find('//') != -1:
                ls = key.split(' ')
                key = ls[len(ls) - 1]
            local_kv_list.append((key.strip(), value.strip()))
        return local_kv_list


def get_cell_style():
    style = xlwt.XFStyle()
    alignment = xlwt.Alignment()
    alignment.horz = 0x01
    alignment.vert = 0x01
    alignment.wrap = 1
    style.alignment = alignment
    return style


def export_excel(workbook):
    file = ''
    if len(sys.argv) > 2:
        file = sys.argv[2]
    if len(file) < 1:
        file = 'web_locales'
    if not file.endswith('.xls'):
        file += '.xls'
    workbook.save('xlsx/'+file)


def main():
    locales_path = 'locales'

    if not os.path.exists(locales_path):
        if len(sys.argv) < 2:
            raise Exception("请在命令行中传入参数 -> 语言包目录")
        web_locales_path = sys.argv[1]
        copy_dir(web_locales_path, locales_path)

    js_dirs = search_files(locales_path, filter_dir)

    workbook = xlwt.Workbook(encoding='utf-8')

    sheets_x = {}
    sheets_y = {}
    sheets = {}

    style = get_cell_style()

    char_width = 256

    for js_dir in js_dirs:
        js_dir_path = locales_path + '/' + js_dir
        js_files = search_files(js_dir_path, filter_js_file)

        for js_file in js_files:
            sheet_name = js_file.replace('.js', '')

            sheets_x_list = sheets_x.get(sheet_name)
            sheets_x_list = sheets_x_list if sheets_x_list else ['key']
            (ix, sheets_x[sheet_name]) = z_arr_push(sheets_x_list, js_dir)

            sheet = sheets.get(sheet_name)
            if not sheet:
                sheet = workbook.add_sheet(sheet_name, cell_overwrite_ok=True)
                sheet.write(0, 0, 'key', style)
                sheet.col(0).width = 7000
                sheets[sheet_name] = sheet

            sheet.col(ix).width = 8000
            sheet.write(0, ix, js_dir, style)

            js_file_path = js_dir_path + '/' + js_file
            arr = read_js_json(js_file_path)

            for (k, v) in arr:
                sheets_y_list = sheets_y.get(sheet_name)
                sheets_y_list = sheets_y_list if sheets_y_list else []
                (iy, sheets_y[sheet_name]) = z_arr_push(sheets_y_list, k)

                sheet.write(iy + 1, 0, k, style)
                sheet.write(iy + 1, ix, v, style)

    export_excel(workbook)


main()
