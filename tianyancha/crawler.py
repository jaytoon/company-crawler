# -*-: coding: utf-8 -*-
"""
:author: albert
:date: 03/08/2019
"""
import logging
from tianyancha.client import TycClient
import xlrd
import xlwt
import os
import time


# from db.mysql_connector import *


def openexcel(file):
    """
    open excel file
    :param file: excel file
    :return: excelojb
    """
    try:
        book = xlrd.open_workbook(file)
        return book
    except Exception as e:
        print("open excel file failed" + str(e))


def readsheets(file):
    """
    read sheet
    :param file: excel obj
    :return: sheet obj
    """
    try:
        book = openexcel(file)
        sheet = book.sheets()
        return sheet
    except Exception as e:
        print("read sheet failed" + str(e))


def readdata(sheet, n=0):
    """
    data read
    :param sheet: excel sheet
    :param n: rows sheet.nrows
    :return: data list
    """
    dataset = []
    for r in range(901, sheet.nrows):
        col = sheet.cell(r, n).value
        # 如果有表头
        if r != 0:
            dataset.append(col)
    return dataset


def tem():
    for sheet in readsheets('测试.xls'):
        for cmyname in readdata(sheet):
            print(cmyname + str(len(cmyname)))
            if len(cmyname) >= 10 and (' ' not in cmyname) and ('、' not in cmyname):
                print(cmyname + str(len(cmyname)))


def start():
    """ 入口函数 """
    g_name_list = []
    g_addr_list = []
    r_name_list = []
    g_desc_list = []
    g_status_list = []
    r_phone_list = []
    r_email_list = []
    id_list = []
    register_capital_list = []
    found_time_list = []
    taxpayer_code_list = []
    register_institute_list = []
    register_code_list = []
    wrong_list = []

    def __printall(elem):
        print(elem.id)
        id_list.append(elem.id)
        register_institute_list.append(elem.register_institute)
        found_time_list.append(elem.found_time)
        register_capital_list.append(elem.register_capital)
        g_name_list.append(elem.name)
        taxpayer_code_list.append(elem.taxpayer_code)
        register_code_list.append(elem.register_code)
        try:
            temp_r_name = elem.representative  # 法定法人
            r_name_list.append(temp_r_name)
        except Exception as e:
            temp_r_name = "暂无法人"
            r_name_list.append(temp_r_name)
            print("没法人")
        try:
            temp_r_phone = elem.contact  # 法人电话
            r_phone_list.append(temp_r_phone)
        except Exception as e:
            temp_r_phone = '暂无信息'
            r_phone_list.append(temp_r_phone)
            print("没电话")
        try:
            temp_g_addr = elem.company_address  # 公司地址
            g_addr_list.append(temp_g_addr)
        except Exception as e:
            temp_g_addr = '暂无信息'
            g_addr_list.append(temp_g_addr)
            print("没地址")
        try:
            temp_g_desc = elem.company_desc  # 公司简介
            g_desc_list.append(temp_g_desc)
        except Exception as e:
            temp_g_desc = '暂无信息'
            g_desc_list.append(temp_g_desc)
            print("没简介")
        try:
            temp_r_email = elem.emails
            r_email_list.append(temp_r_email)
        except Exception as e:
            temp_r_email = '暂无信息'
            print("没邮箱")
            r_email_list.append(temp_r_email)
        temp_g_status = elem.biz_status  # 企业状态
        g_status_list.append(temp_g_status)
        logging.info(elem.__str__())

    keys = globals().get('keywords', [])
    for key in keys:
        logging.info('正在采集[%s]...' % key)
        companies = TycClient().search(key).companies
        # 写入db
        # insert_company(companies)
        try:
            company = companies[0]
            __printall(company)
        except Exception as e:
            wrong_list.append(key)
            print(repr(e))
    logging.info("completed")
    workbook = xlwt.Workbook()
    # 创建sheet对象，新建sheet
    sheet1 = workbook.add_sheet('天眼查数据', cell_overwrite_ok=True)
    sheet2 = workbook.add_sheet('未查数据', cell_overwrite_ok=True)
    # ---设置excel样式---
    # 初始化样式
    style = xlwt.XFStyle()
    # 创建字体样式
    font = xlwt.Font()
    font.name = '仿宋'
    #    font.bold = True #加粗
    # 设置字体
    style.font = font
    # 使用样式写入数据
    print('正在存储数据，请勿打开excel')
    # 向sheet中写入数据
    name_list = ['公司ID', '公司名称', '公司地址', '法定法人', '注册资本', '注册时间', '公司状态', '法人邮箱', '法人电话', '纳税人识别号',
                 '登记机关', '工商注册号', '公司简介']
    for cc in range(0, len(name_list)):
        sheet1.write(0, cc, name_list[cc], style)
    for dd in range(0, len(wrong_list)):
        sheet2.write(0, dd, wrong_list[dd], style)
    for i in range(0, len(g_name_list)):
        print(g_name_list[i])
        sheet1.write(i + 1, 0, id_list[i], style)  # 公司ID
        sheet1.write(i + 1, 1, g_name_list[i], style)  # 公司名字
        sheet1.write(i + 1, 2, g_addr_list[i], style)  # 公司地址
        sheet1.write(i + 1, 3, r_name_list[i], style)  # 法定法人
        sheet1.write(i + 1, 4, register_capital_list[i], style)  # 注册资本
        sheet1.write(i + 1, 5, found_time_list[i], style)  # 成立日期
        sheet1.write(i + 1, 6, g_status_list[i], style)  # 公司状态
        sheet1.write(i + 1, 7, r_email_list[i], style)  # 法人邮箱
        sheet1.write(i + 1, 8, r_phone_list[i], style)  # 法人电话
        sheet1.write(i + 1, 9, taxpayer_code_list[i], style)  # 纳税人识别号
        sheet1.write(i + 1, 10, register_institute_list[i], style)  # 登记机关
        sheet1.write(i + 1, 11, register_code_list[i], style)  # 工商注册号
        sheet1.write(i + 1, 12, g_desc_list[i], style)  # 公司简介
    for i in range(0, len(wrong_list)):
        sheet2.write(i + 1, 0, wrong_list[i], style)
    # 保存excel文件，有同名的直接覆盖
    workbook.save(os.getcwd() + r'\\' + time.strftime("%Y-%m-%d-%H-%M-%S", time.localtime()) + ".xls")
    print('保存完毕~')


def load_keys(keys: list):
    globals().setdefault('keywords', keys)
