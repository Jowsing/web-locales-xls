# -*- coding: UTF-8 -*-

import sys
import os
import xlrd
import xlwt
import shutil


def z_log(value):
    print(value)
    print('\n')


def r_newline(str):
    return str.replace('\n', '').replace('\r', '')


def lines_in_js(str):
    str_insert_line_char = str.replace(
        '\',\n', '\'<--line-->').replace('",\n', '"<--line-->').replace('`,\n', '`<--line-->')
    return r_newline(str_insert_line_char).split('<--line-->')


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
        js_str = js_str[js_str.find('{') + 1: js_str.rfind('}')]  # 去除 {}
        local_str_list = lines_in_js(js_str)  # 插入行标，并去除换行
        local_kv_list = []
        for local_str in local_str_list:
            if len(local_str) == 0:
                continue
            str = r_newline(local_str.strip())
            kvArr = str.split(':', 1)
            [key, value] = kvArr
            if key.find('//') != -1:
                ls = key.split(' ')
                key = ls[len(ls) - 1]
            local_kv_list.append((key.strip(), value.strip()))
        return local_kv_list


def get_cell_style():
    style = xlwt.XFStyle()
    alignment = xlwt.Alignment()
    alignment.horz = 0x02
    alignment.vert = 0x01
    alignment.wrap = 1
    style.alignment = alignment
    return style


def import_excel(file_path):
    book = xlrd.open_workbook(file_path)
    return book


def export_excel(workbook):
    file = ''
    if len(sys.argv) > 2:
        file = sys.argv[2]
    if len(file) < 1:
        file = 'web_locales'
    if not file.endswith('.xls'):
        file += '.xls'
    workbook.save('xlsx/'+file)


def make_dir(path):
    if not os.path.exists(path):
        os.makedirs(path)


def delete_dir(path):
    if os.path.exists(path):
        shutil.rmtree(path)


def write_js_file(path, str):
    if os.path.exists(path):
        os.remove(path)
    file = open(path, 'a')
    file.write(str)
    file.close()


def js_objstr_add(str, key, value):
    if len(str) < 1:
        str += 'export default {'
    str += '\n  {0}: {1},'.format(key.encode(encoding='utf-8'),
                                  value.encode(encoding='utf-8'))
    return str


def get_sys_arg(i, error_msg='', def_value=''):
    if len(sys.argv) > i:
        return sys.argv[i]
    elif len(def_value) > 0:
        return def_value
    else:
        raise Exception("请在命令行中传入第 {0} 个参数 -> {1}".format(i, error_msg))
