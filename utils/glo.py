#!/usr/bin/env python
# encoding: utf-8
class Globals:
    def __init__(self) -> None:
        #初始化一个全局的字典
        pass

    def _init(self):
        global _global_dict
        _global_dict = {}

    def set_value(self, key, value):
        _global_dict[key] = value
        
    def get_value(self, key):
        try:
            return _global_dict[key]
        except KeyError as e:
            e