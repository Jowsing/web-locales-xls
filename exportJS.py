# -*- coding: UTF-8 -*-

from utils import *


def main():
    locales_path = get_sys_arg(2, '', 'locales_new')

    execl_file_path = get_sys_arg(1, '语言包的 Excel 文件')

    book = import_excel(execl_file_path)

    sheet_len = book.nsheets

    if sheet_len > 0:
        make_dir(locales_path)

    for sheet_i in range(sheet_len):
        sheet = book.sheet_by_index(sheet_i)
        rows = sheet.nrows
        cols = sheet.ncols
        file_name = sheet.name
        for col in range(cols):
            if col < 1:
                continue

            dir = locales_path + '/' + sheet.cell_value(0, col)
            make_dir(dir)

            js = ''
            for row in range(rows):
                if row < 1:
                    continue
                key = sheet.cell_value(row, 0)
                value = sheet.cell_value(row, col)
                js = js_objstr_add(js, add_first_last(
                    key), add_first_last(value))

            if len(js) < 1:
                continue

            js += '\n};\n'

            file_path = dir + '/' + file_name + '.js'
            write_js_file(file_path, js)


main()
