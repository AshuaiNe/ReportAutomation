#!/usr/bin/env python
# encoding: utf-8
import time
import shutil
from utils.glo import Globals
from pathlib import Path


class ExistsMkDir:
    def __init__(self) -> None:
        self.glo = Globals()
        self._path = self.glo.get_value('_path')
        self.parsing = self.glo.get_value('parsing')
        self._result_file = f"{self._path}/compare/result/{self.parsing}/{time.strftime('%Y-%m-%d', time.localtime())}/"
        self._excel_report = f"{self._path}/compare/excel_report/{self.parsing}/{time.strftime('%Y-%m-%d', time.localtime())}/"
        self._challenger = f"{self._path}/compare/challenger/{self.parsing}/{time.strftime('%Y-%m-%d', time.localtime())}/"
        self._original_data = f"{self._path}/compare/original_data/{self.parsing}/{time.strftime('%Y-%m-%d', time.localtime())}/"
        self.path_lists = [self._excel_report, self._result_file, self._challenger, self._original_data]

    def exists_mk_dir(self, bol=True):
        for x in self.path_lists:
            if x in (f"{self._result_file}", f"{self._excel_report}") and bol:
                if Path(x).exists():
                    shutil.rmtree(x, ignore_errors=True)
                Path(x).mkdir()
            else:
                if Path(x).exists():
                    continue
                else:
                    Path(x).mkdir()