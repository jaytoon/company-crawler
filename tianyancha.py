#!/usr/bin/python3
# -*-: coding: utf-8 -*-
"""
:author: lubosin
:date: 03/28/2019
"""
from tianyancha import crawler
from util import log
import urllib3
import xlwt
import os
import time
import re

urllib3.disable_warnings()

log.set_file("./logs/tianyancha.log")


def tem(keys: list):
    for sheet in crawler.readsheets('企业去重数据.xls'):
        for cmyname in crawler.readdata(sheet):
            # print(cmyname + str(len(cmyname)))
            if len(cmyname) >= 7 and (' ' not in cmyname) and ('、' not in cmyname) and ('，' not in cmyname) and (
                    re.match(r"[\u4e00-\u9fa5]{2,3}[（\(][\u4e00-\u9fa5]+[）\)]", cmyname) is None):
                keys.append(cmyname)
                print(cmyname + str(len(cmyname)))


if __name__ == '__main__':
    # keys = ['大连阿大海产养殖有限公司']
    keys = []
    tem(keys)
    keys = set(keys)
    crawler.load_keys(keys)

    crawler.start()  # 东港市种畜场 旅顺千品渔港 盘山县水利局 营口市水利局 东港市水利局 大连土城盐场
