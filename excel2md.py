#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import xlwings as xw

def mkline(datas):
    line_str = '| '
    for data in datas:
        if isinstance(data, str):
            line_str += data
        else:
            line_str += ' '
        line_str += ' | '

    line_str += '\n'
    return line_str

def mk_header(datas_format):
    line_str = '| '
    for data in datas_format:
        line_str += '---'
        line_str += ' | '

    line_str += '\n'
    return line_str


if __name__ == '__main__':
    while True:
        file_name = input("Please input excel name with extensions: ")
        print("your input file is :[" + file_name +"]")

        try:
            wb = xw.Book(file_name)
        except BaseException:
            print("Sorry, I can't open this file, please retry")
            continue

        break

    for sheet in wb.sheets:
        sheet.activate() # active current sheet
        sheet_info = sheet.used_range
        row_max = sheet_info.last_cell.row #last row number
        column_max = sheet_info.last_cell.column #last column number
        if row_max <= 1 and column_max <= 1:
            continue

        path = './' + sheet.name + '.md'
        with open(path, 'w+') as w_file:
            for row_index in range(0, row_max):
                index_start = sheet.cells[row_index, 0]
                index_end = sheet.cells[row_index, column_max-1]
                line_values = sheet.range(index_start, index_end).value
                w_file.write(mkline(line_values)) 

                if row_index == 0:
                    w_file.write(mk_header(line_values)) 

        w_file.close()

    wb.close()
    print("file is output, the exe was end")
