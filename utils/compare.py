#!/usr/bin/env python
# encoding: utf-8
import re
from pathlib import Path
from deepdiff import DeepDiff
from typing import List
from utils.parsing_word_document import ParsingWord
from utils.parsing_excel import ParingExcel
from utils.parsing_txt import ParsingTxt


class Compare:

    def __init__(self) -> None:
        self.parsing_word = ParsingWord()
        self.parsing_excel = ParingExcel()
        self.parsing_txt = ParsingTxt()
        self._pattern = re.compile(r"(?<=\[).*?(?=\])")

    def compare_deepdiff_excel(self, begin_filename, end_filename):
        '''
        @description: 获取两份excel文档差异性
        @param {begin_filename, end_filename}
        @return {json}
        '''  
        added_sheet = 0
        removed_sheet = 0
        sheet_cell = 0
        list_args: List = [begin_filename, end_filename]
        compare_list: List = []
        for args in list_args:
            compare_list.append(self.parsing_excel.get_excel(args))
        result = DeepDiff(compare_list[0], compare_list[1])
        self.parsing_excel.set_excel(list_args[1], result)
        for key in result.keys():
            if 'dictionary_item_added' == key:
                added_sheet = len(result[key])
            elif 'dictionary_item_removed' == key:
                removed_sheet = len(result[key])
            elif 'values_changed' == key:
                sheet_cell = len(result[key])
        result = [begin_filename, end_filename.replace('result', 'challenger'), end_filename, added_sheet, removed_sheet,
                    sheet_cell, False if result else True]
        return result

    def compare_deepdiff_txt(self, begin_filename, end_filename):
        '''
        @description: 获取两份txt文档差异性
        @param {begin_filename, end_filename}
        @return {json}
        '''  
        added_line = 0
        removed_line = 0
        line = 0
        list_args: List = [begin_filename, end_filename]
        compare_list: List = []
        for args in list_args:
            compare_list.append(self.parsing_txt.get_txt(args))
        result = DeepDiff(compare_list[0], compare_list[1])
        begin_filename = str(begin_filename).split('.')[0] + '.xls'
        end_filename = str(end_filename).split('.')[0] + '.xls'
        self.parsing_txt.set_json_to_excel(end_filename, result)
        for key in result.keys():
            if 'dictionary_item_added' == key:
                added_line = len(result[key])
            elif 'dictionary_item_removed' == key:
                removed_line = len(result[key])
            elif 'values_changed' == key:
                line = len(result[key])
        result = [begin_filename, end_filename.replace('result', 'challenger'), end_filename, added_line, removed_line,
                    line, False if result else True]
        return result

    def compare_deepdiff_word(self, begin_filename, end_filename) -> str:
        '''
        @description: 获取两份word文档差异性
        @param {begin_filename, end_filename}
        @return {json}
        '''    
        list_args: List = [begin_filename, end_filename]
        compare_list: List = []
        table_list = []
        dictionary_item_added = 0
        docx_paragraphs_removed = 0
        docx_tables_removed = 0
        docx_paragraphs = 0
        docx_tables = 0
        docx_header = 0
        for args in list_args:
            compare_data = {}
            compare_data['docx_paragraphs'], compare_data['docx_tables'], compare_data['docx_header'] = self.parsing_word.get_word(args)
            compare_list.append(compare_data)
        result = DeepDiff(compare_list[0], compare_list[1])
        for x in compare_list[1]['docx_tables'].keys():
            table_list.append(x)
        self.parsing_word.set_docx(list_args[1], table_list, result)
        for key in result.keys():
            if 'dictionary_item_added' == key:
                dictionary_item_added = len(result[key])
            elif 'dictionary_item_removed' == key:
                for x in result[key]:
                    self.dictionary_item_removed = len(result[key])
                    find_str = re.findall(r"\[.*?\]", x)
                    paragraphs_or_tables = find_str[0].replace("[", "").replace("]", "").replace("'", "")
                    if 'docx_paragraphs' == paragraphs_or_tables:
                        docx_paragraphs_removed += 1
                    elif 'docx_tables' == paragraphs_or_tables:
                        docx_tables_removed += 1
            elif 'values_changed' == key:
                for x in result[key]:
                    find_str = re.findall(r"\[.*?\]", x)
                    paragraphs_or_tables_or_header = find_str[0].replace("[", "").replace("]", "").replace("'", "")
                    if 'docx_paragraphs' == paragraphs_or_tables_or_header:
                        docx_paragraphs += 1
                    elif 'docx_tables' == paragraphs_or_tables_or_header:
                        docx_tables += 1
                    elif 'docx_header' == paragraphs_or_tables_or_header:
                        docx_header += 1
        result = [begin_filename, end_filename.replace('result', 'challenger'), end_filename, dictionary_item_added, docx_paragraphs_removed,
                docx_tables_removed, docx_paragraphs, docx_tables, docx_header, False if result else True]
        return result
