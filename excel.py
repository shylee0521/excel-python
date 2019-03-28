#!/usr/bin/env python
# -*- coding: utf-8 -*-

# @Date    : 2019/3/26
# @Author  : liyan (li.yan@eisoo.com)

import datetime
import xlrd
import xlutils.copy
import xlwt
import os


class ExcelFactory(object):
    def get_excel_handler(self, filename):
        if os.path.exists(filename):
            return AppendExcel(filename)
        else:
            return NewExcel(filename)


class BaseExcel(object):
    def __init__(self, filename):
        self.excel_handler = None
        self.file_name = filename

    def add_sheet(self, sheet_name, _header):
        sheet_handler = self.excel_handler.add_sheet(sheet_name)
        for col in range(4):
            col_handler = sheet_handler.col(col)
            col_handler.width = 256 * 20
        self.write_header(sheet_handler, 0, _header)
        return sheet_handler

    def write_header(self, sheet_handler, row_num, line_value):
        style = xlwt.XFStyle()
        font = xlwt.Font()
        font.bold = True
        style.font = font
        self.write_line(sheet_handler, row_num, line_value, style)

    def write_line(self, sheet_handler, row_num, line_value, font=None):
        col_num = 0
        for value in line_value:
            if isinstance(value, (datetime.date, datetime.datetime)):
                value = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            if font is None:
                sheet_handler.write(row_num, col_num, value)
            else:
                sheet_handler.write(row_num, col_num, value, font)
            col_num += 1

    def write_data(self, src_data_list, header_list, sheet_handler=None, start_row_num=0, sheet_name='sheet'):
        for index, item_data in enumerate(src_data_list):
            sheet_name = sheet_name + str(index + 1)
            if start_row_num == 0:
                start_row_num = 1
                # 没有记录时，也要把header写入
                sheet_handler = self.add_sheet(sheet_name, header_list)
            for line in item_data:
                if start_row_num == 0:
                    sheet_handler = self.add_sheet(sheet_name, header_list)
                    start_row_num += 1
                self.write_line(sheet_handler, start_row_num, line)
                start_row_num += 1
                if start_row_num > 65535:
                    start_row_num = 0
        self.excel_handler.save(self.filename)


class AppendExcel(BaseExcel):
    def __init__(self, filename):
        self.filename = filename
        self.excel_reader = None
        self.excel_handler = self.get_excel_handler(filename)

    def get_excel_handler(self, file_name):
        self.excel_reader = xlrd.open_workbook(file_name, formatting_info=True)
        return xlutils.copy.copy(self.excel_reader)

    def get_all_sheets(self):
        return self.excel_reader.sheet_names()

    def sheet_already_rows_by_sheet_name(self, sheet_name):
        sheet_handler = self.excel_reader.sheet_by_name(sheet_name)
        return sheet_handler.nrows

    def build_excel(self, src_data_list, header_list):
        sheets = self.get_all_sheets()
        latest_sheet = sheets[-1]
        sheet_handler = self.excel_handler.get_sheet(len(sheets) - 1)
        rows = self.sheet_already_rows_by_sheet_name(latest_sheet)
        if rows == 65535:
            rows = 0

        self.write_data(src_data_list, header_list, sheet_handler, rows, latest_sheet)


class NewExcel(BaseExcel):
    def __init__(self, filename):
        self.filename = filename
        self.excel_handler = self.get_excel_handler()

    def get_excel_handler(self, encoding='utf-8'):
        return xlwt.Workbook(encoding=encoding)

    def build_excel(self, src_data_list, header_list):
        """
            构建excel文件或者写入流对象
        @param src_data_list: 源数据列表
        @param header_list: 表头
        @param sheet_name_list: 单元表
        @return:
        """
        self.write_data(src_data_list, header_list)


if __name__ == '__main__':
    text = ['a', 'b', 'c', 'd', 'e']
    excel_filename = 'excel_%s.xls' % datetime.datetime.now().strftime('%Y-%m-%d-%H-%M-%S')
    excel_handler = ExcelFactory().get_excel_handler(excel_filename)
    excel_handler.build_excel([text], ['a'])
