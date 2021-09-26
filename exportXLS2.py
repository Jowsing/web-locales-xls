# -*- coding: UTF-8 -*-

from utils import *

root_path = sys.path[0]
locales_path = 'locales'
char_width = 256


def main():
    if not os.path.exists(locales_path):
        web_locales_path = get_sys_arg(1, '语言包目录')
        copy_dir(web_locales_path, locales_path)

    js_files = search_files(locales_path, filter_js_file)
    js_files = sort_js_by_langs(js_files)

    workbook = xlwt.Workbook(encoding='utf-8')
    sheet = workbook.add_sheet('语言包', cell_overwrite_ok=True)

    sheet.col(0).width = char_width * 35

    col_index = 1
    row_keys = []

    style = get_cell_style()

    for js_file in js_files:
        col_name = js_file.replace('.js', '')
        arr = read_js_json(locales_path + '/' + js_file)

        sheet.col(col_index).width = char_width * 35
        sheet.write(0, col_index, col_name, style)

        for (k, v) in arr:
            safe_append(row_keys, k)

            row_index = len(row_keys)

            if len(k) > 0:
                sheet.write(row_index, 0, replace_first_last(k, '\''), style)
                sheet.write(row_index, col_index,
                            replace_first_last(v, '\''), style)
            else:
                sheet.write(row_index, 0, k, style)
                sheet.write(row_index, col_index, v, style)

        col_index += 1

    export_excel(workbook)
    delete_dir(locales_path)


main()
