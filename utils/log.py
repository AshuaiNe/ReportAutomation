#!/usr/bin/env python
# encoding: utf-8
from enum import Flag
import time
import logging
import inspect
from pathlib import Path
 
 
class LoggerFactory:
    level_relations = {
        'debug': logging.DEBUG, 'info': logging.INFO, 'warning': logging.WARNING,
        'error': logging.ERROR, 'critical': logging.CRITICAL
    }

    def __init__(self, level='info'):
        """
        实例化LoggerFactory类时的构造函数
        :param name: 
        """
        # 实例化logging
        self.name = inspect.stack()[1].function
        self.logger = logging.getLogger(self.name)
        # 输出的日志格式
        self.formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        self.level = self.level_relations.get(level)
 
    def create_logger(self):
        """
        构造一个日志对象
        :return:
        """
        # 设置日志级别
        self.logger.setLevel(self.level)
        # 设置日志输出的文件
        self.directory = Path.joinpath(Path(__file__).parent.parent, 'log')
        handle = logging.FileHandler(f"{self.directory}/{time.strftime('%Y-%m-%d', time.localtime())}.log", encoding='utf-8')
        # 输出到日志文件的日志级别
        handle.setLevel(self.level)
        handle.setFormatter(self.formatter)
        self.logger.addHandler(handle)
        # 输出到控制台的显示信息
        console = logging.StreamHandler()
        console.setLevel(self.level)
        console.setFormatter(self.formatter)
        self.logger.addHandler(console)
