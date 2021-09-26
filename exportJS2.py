# -*- coding: UTF-8 -*-
from utils import *


def main():
    locales_path = get_sys_arg(2, '', 'locales_new')

    execl_file_path = get_sys_arg(1, '语言包的 Excel 文件')

    book = import_excel(execl_file_path)

    if book.nsheets == 0:
        return

    make_dir(locales_path)

    sheet = book.sheet_by_index(0)
    rows = sheet.nrows
    cols = sheet.ncols
    for col in range(cols):
        if col < 1:
            continue

        file_name = sheet.cell_value(0, col)

        if len(file_name) == 0:
            continue

        js = ''
        for row in range(rows):
            if row < 1:
                continue
            key = sheet.cell_value(row, 0)
            value = sheet.cell_value(row, col)
            js = js_objstr_add(js, add_first_last(key),
                               add_first_last(value))

            if len(js) < 1:
                continue
        js += '\n};\n'

        file_path = locales_path + '/' + file_name + '.js'
        write_js_file(file_path, js)


main()
