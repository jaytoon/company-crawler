#!/usr/bin/python3
# -*-: coding: utf-8 -*-
"""
:author: albert
:date: 03/08/2019
"""
import logging
from tianyancha.client import TycClient
# from db.mysql_connector import *


def start():
    """ 入口函数 """
    def __printall(items):
        for elem in items:
            logging.info(elem.__str__())

    keys = globals().get('keywords', [])
    for key in keys:
        logging.info('正在采集[%s]...' % key)
        companies = TycClient().search(key).companies
        # 写入db
        # insert_company(companies)
        __printall(companies)
    logging.info("completed")


def load_keys(keys: list):
    globals().setdefault('keywords', keys)





