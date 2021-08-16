#!/usr/bin/env python
# encoding: utf-8
import requests
import re


class Message:

    def __init__(self) -> None:
        self.dictionary_item_added = 0
        self.dictionary_item_removed = 0
        self.docx_paragraphs = 0
        self.docx_tables = 0
        self.docx_header = 0
    
    def tx_message(self, msg):
        if not msg:
            text_json = {
                "msgtype": "markdown",
                "markdown": {
                    "content": "word文档比对结果，请相关同事注意。\n>无差异部分: "
                }
            }
        else:
            text_json = {
                "msgtype": "markdown",
                "markdown": {
                    "content": f"word文档比对结果，请相关同事注意。\n>比对文件数:{msg[0]}<font color=\"comment\"></font>\n>新增章节:{msg[4]}<font color=\"comment\"></font>\n>删除章节:{msg[5]} <font color=\"comment\"></font>\n>删除表数量:{msg[6]} <font color=\"comment\"></font>\n>差异段落:{msg[7]} <font color=\"comment\"></font>\n>差异单元格:{msg[8]} <font color=\"comment\"></font>\n>差异页眉:{msg[9]} <font color=\"comment\"></font>"
                }
            }
        requests.post("https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=4166a9b2-bc36-4f39-ba44-ab8a435fe4c3", headers={"Content-Type": "application/json"}, json=text_json).json()
