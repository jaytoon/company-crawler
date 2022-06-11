import os
from tianyancha import crawler
import re

import xlwt
if __name__ == '__main__':
    key_list = []
    for sheet in crawler.readsheets('用海企业.xls'):
        for cmyname in crawler.readdata(sheet):
            # print(cmyname + str(len(cmyname)))
            if len(cmyname) >= 7 and (' ' not in cmyname) and ('、' not in cmyname) and ('，' not in cmyname) and (
                    re.match(r"[\u4e00-\u9fa5]{2,3}[（\(][\u4e00-\u9fa5]+[）\)]", cmyname) is None):
                key_list.append(cmyname)
                # print(cmyname + str(len(cmyname)))
    print(len(key_list))
    key_list=set(key_list)
    key_list=list(key_list)

    print(len(key_list))
    workbook = xlwt.Workbook()
    # 创建sheet对象，新建sheet
    sheet1 = workbook.add_sheet('海岛数据', cell_overwrite_ok=True)
    # ---设置excel样式---
    # 初始化样式
    style = xlwt.XFStyle()
    # 创建字体样式
    font = xlwt.Font()
    font.name = '宋体'
    #    font.bold = True #加粗
    # 设置字体
    style.font = font
    # 使用样式写入数据
    print('正在存储数据，请勿打开excel')
    # 向sheet中写入数据
    title_list = ['使用人名称']
    for cc in range(0, len(title_list)):
        sheet1.write(0, cc, title_list[cc], style)
    for i in range(0, len(key_list)):
        print(key_list[i])
        sheet1.write(i + 1, 0, key_list[i], style)  # 海岛名称
        # sheet1.write(i + 1, 1, loc_list[i], style)  # 经纬度
        # sheet1.write(i + 1, 2, g_addr_list[i], style)  # 公司地址
        # sheet1.write(i + 1, 3, r_name_list[i], style)  # 法定法人
        # sheet1.write(i + 1, 4, register_capital_list[i], style)  # 注册资本
        # sheet1.write(i + 1, 5, found_time_list[i], style)  # 成立日期
        # sheet1.write(i + 1, 6, g_status_list[i], style)  # 公司状态
        # sheet1.write(i + 1, 7, r_email_list[i], style)  # 法人邮箱
        # sheet1.write(i + 1, 8, r_phone_list[i], style)  # 法人电话
        # sheet1.write(i + 1, 9, taxpayer_code_list[i], style)  # 纳税人识别号
        # sheet1.write(i + 1, 10, register_institute_list[i], style)  # 登记机关
        # sheet1.write(i + 1, 11, register_code_list[i], style)  # 工商注册号
        # sheet1.write(i + 1, 12, g_desc_list[i], style)  # 公司简介
    # 保存excel文件，有同名的直接覆盖
    workbook.save(os.getcwd() + r"\\企业去重数据.xls")
    print('保存完毕~')
