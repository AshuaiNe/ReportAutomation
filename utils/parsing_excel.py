#!/usr/bin/env python
# encoding: utf-8
import re
import pandas
import time
import xlwt
import string
from pathlib import Path
from utils.glo import Globals
from win32com import client as wc
from openpyxl import load_workbook
from openpyxl.comments import Comment


class ParingExcel:

    def __init__(self) -> None:
        self.glo = Globals()
        self._path = self.glo.get_value('_path')
        self.parsing = self.glo.get_value('parsing')
        self.product = self.glo.get_value('product')
        self._excel_report = f"{self._path}/compare/excel_report/{self.parsing}/{time.strftime('%Y-%m-%d', time.localtime())}/"

    def write_excel(self, data_lists, app):
        for x in range(len(data_lists)):
            for y in range(3):
                if y == 1:
                    filename = Path(data_lists[x][y]).name.replace("copyfile", "")
                    filepath = data_lists[x][y].replace("copyfile", "")
                else:
                    filename = Path(data_lists[x][y]).name
                    filepath = data_lists[x][y]
                data_lists[x][y] = f'=HYPERLINK("{filepath}", "{filename}")'
        if app == 'word':
            columns=['原始报告', '最新报告', '结果文件', '新增章节数', '删除章节数', '删除表数量', '段落差异数', '单元格差异数', '页眉差异数', '结果']
        elif app == 'excel':
            columns=['原始报告', '最新报告', '结果文件', '新增sheet数', '删除sheet数', '单元格差异', '结果']
        elif app == 'txt':
            columns=['原始报告', '最新报告', '结果文件', '新增行', '删除行', '行数据差异', '结果']
        df = pandas.DataFrame(data_lists, columns=columns)
        df.to_excel(f"{self._excel_report}{time.strftime('%Y-%m-%d', time.localtime())}.xlsx", index=None, sheet_name="测试报告")

    def get_excel(self, filename):
        sheet_dict = {}
        alphabet = list(string.ascii_uppercase) + [letter1+letter2 for letter1 in string.ascii_uppercase for letter2 in string.ascii_uppercase]
        sheet_names = pandas.ExcelFile(filename, engine='xlrd').sheet_names
        for sheet in sheet_names:
            sheet_json = {}
            cell_sum = 0
            cell_list = []
            dr = pandas.read_excel(filename, engine='xlrd', sheet_name=sheet)
            for key in dr.to_dict().keys():
                cell_list = ['None' if str(x) == 'nan' else str(x) for x in dr.to_dict()[key].values()]
                cell_list.insert(0, key)
                sheet_json[alphabet[cell_sum]] = cell_list
                cell_sum += 1
            sheet_dict[sheet] = sheet_json
        return sheet_dict

    def set_excel(self, filename, compare_result):
        added_sheet = []
        removed_sheet = []
        changed = []
        values_changed = []
        for key in compare_result.keys():
            if key == 'dictionary_item_added':
                for x in range(len(compare_result[key])):
                    find_str = re.findall(r"\'.*?\'", compare_result[key][x])
                    y = find_str[0].replace("'", "")
                    added_sheet.append(f'{y}')
            elif key == 'dictionary_item_removed':
                for x in range(len(compare_result[key])):
                    find_str = re.findall(r"\'.*?\'", compare_result[key][x])
                    y = find_str[0].replace("'", "")
                    removed_sheet.append(f'{y}')
            elif key == 'values_changed':
                for key_1 in compare_result[key]:
                    find_str = re.findall(r"\[.*?\]", key_1)
                    find_str_0 = find_str[0].replace("[", "").replace("]", "").replace("'", "")
                    find_str_1 = find_str[1].replace("[", "").replace("]", "").replace("'", "")
                    find_str_2 = int(find_str[2].replace("[", "").replace("]", "").replace("'", "")) + 1
                    changed.append(f'{find_str_0}：{find_str_1}{find_str_2}')
                    values_changed.append(compare_result[key][key_1])
            else:
                continue
        dict_frame = {"新增sheet": added_sheet, "删除sheet": removed_sheet, "差异单元格": changed, "差异结果": values_changed}
        wb = xlwt.Workbook(encoding='utf-8')
        ws = wb.add_sheet('比对结果', cell_overwrite_ok=True)
        for x, key in enumerate(dict_frame):
            ws.write(0, x, str(key))
            row = 1
            for y in range(len(dict_frame[key])):
                ws.write(row, x, str(dict_frame[key][y]))
                row += 1
        wb.save(filename)
