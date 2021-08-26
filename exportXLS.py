# -*- coding: UTF-8 -*-

from utils import *

root_path = sys.path[0]


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
                sheet.col(0).width = char_width * 25
                sheets[sheet_name] = sheet

            sheet.col(ix).width = char_width * 35
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
    delete_dir(locales_path)


main()
