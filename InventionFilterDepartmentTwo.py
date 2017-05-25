#!/bin/usr/env python
# -*- coding:cp936 -*-
"""这个工具的作用是过滤部门的专利总览表，然后统计出一和二处个人的专利完成情况。分为六列：组长名、组员名、发明提交、发明受理、实用新型提交、实用新型受理。输入文档
必须是csv格式的，且只有一个sheet。输出文件名为《测试验证部测试%s个人专利完成情况统计.csv》,依据不同的处别而不同。 本工具基于Python Tkinter制作图形界面。依赖详见import部分。
打包成exe格式请使用pyinstall,命令为Python pyinstaller.py -F InventionFilterDepartmentOne.py
"""

import numpy
import pandas
import Tkinter
import tkMessageBox
import ttk
import tkFileDialog
import os
import time
import xlsxwriter
import xlrd


filename_original = unicode()
dir_filename_display = unicode()
root = Tkinter.Tk()
root.title("专利结果过滤工具".decode('gbk'))
# root.geometry('800x600')
root.resizable(width=True, height=True)
var_char_entry_filename_need_filter = Tkinter.StringVar()
var_char_combox_department = Tkinter.StringVar()
var_char_entry_filename_after_filter = Tkinter.StringVar()

TeamLeader = ['贾岛'.decode('gbk'), '潘霖'.decode('gbk'), '韩琳琳'.decode('gbk'), '苗永威'.decode('gbk'),
              '史沛玉'.decode('gbk'), '杨文清'.decode('gbk'), '伯绍文'.decode('gbk'), '迟江波'.decode('gbk'), '李永亮'.decode('gbk'),
              '曹翔'.decode('gbk')]
# 万浩离职
Member0 = ['贾岛'.decode('gbk'), '李光达'.decode('gbk'), '刘茂峰'.decode('gbk'), '范鹏飞'.decode('gbk'), '谭静静'.decode('gbk'),
           '张文珂'.decode('gbk'), '代如静'.decode('gbk')]
Member1 = ['潘霖'.decode('gbk'), '刘博'.decode('gbk'), '黄翼'.decode('gbk'), '董喜燕'.decode('gbk'), '郝良晟'.decode('gbk')]
Member2 = ['韩琳琳'.decode('gbk'), '林海'.decode('gbk'), '曹加峰'.decode('gbk'), '张行武'.decode('gbk'), '李建波'.decode('gbk'),
           '冯晓洁'.decode('gbk')]
Member3 = ['苗永威'.decode('gbk'), '闫硕'.decode('gbk'), '刘瑞雪'.decode('gbk'), '王云鹏'.decode('gbk'), '卢正超'.decode('gbk')]
Member4 = ['史沛玉'.decode('gbk'), '张超'.decode('gbk'), '张锟'.decode('gbk'), '刘智刚'.decode('gbk'), '巩祥文'.decode('gbk'),
           '孙玉超'.decode('gbk'), '韩超'.decode('gbk'), '徐伟超'.decode('gbk'), '赵盛'.decode('gbk'), '王建刚'.decode('gbk'),
           '高莹'.decode('gbk'), '王旭林'.decode('gbk'), '杨惠'.decode('gbk'), '程佳佳'.decode('gbk'),]
Member5 = ['杨文清'.decode('gbk'), '贠雄斌'.decode('gbk'), '孙薇'.decode('gbk'), '李静'.decode('gbk'), '杨永峰'.decode('gbk')]
Member6 = ['伯绍文'.decode('gbk'), '李波'.decode('gbk'), '刘东伟'.decode('gbk'), '吴培琴'.decode('gbk'), '武秋星'.decode('gbk'),
           '胥志泉'.decode('gbk'), '赵召'.decode('gbk'), '李壮'.decode('gbk'), '李俊卿'.decode('gbk')]
Member7 = ['迟江波'.decode('gbk'), '刘浩君'.decode('gbk'), '李彦华'.decode('gbk'), '韩燕燕'.decode('gbk'),
            '梁恒勋'.decode('gbk'), '黄锦盛'.decode('gbk')]
Member8 = ['李永亮'.decode('gbk'), '李丹6011'.decode('gbk'), '兰太顺'.decode('gbk')]
Member9 = ['曹翔'.decode('gbk'), '康艳丽'.decode('gbk'), '王智仙'.decode('gbk'), '邓振宏'.decode('gbk'), '蔣弦佑'.decode('gbk')]

TitleItem = ['组长名'.decode('gbk'), '组员名'.decode('gbk'), '发明受理数量'.decode('gbk'), '发明提交数量'.decode('gbk'),
             '实用新型受理数量'.decode('gbk'), '实用新型提交数量'.decode('gbk')]


def get_filename():
    global filename_original
    filename_iometer = tkFileDialog.askopenfilename()
    var_char_entry_filename_need_filter.set(filename_iometer)
    filename_original = filename_iometer


def set_filename():
    global dir_filename_display
    dir_filename_display = tkFileDialog.askdirectory().replace('/', '\\')
    var_char_entry_filename_after_filter.set(dir_filename_display)
#    filename_display = os.path.join(dir_filter_iometer, "测试验证部测试%s个人专利完成情况统计.csv".decode('gbk') )


def get_data():
    length_team = len(TeamLeader)
    filename_input = filename_original
    department_to_filter = var_char_combox_department.get()
    timestamp = time.strftime('%Y%m%d', time.localtime())
    filename_output = os.path.join(dir_filename_display,
                                   "%s个人专利完成情况统计-%s.xlsx".decode('gbk') % (department_to_filter, timestamp))
    WorkBook = xlsxwriter.Workbook(filename_output)
    SheetOne = WorkBook.add_worksheet('各组专利完成情况统计'.decode('gbk'))
    format = WorkBook.add_format()
    format.set_border(1)
    sum_line = 0
    ListUsername = []
    for i in range(0, length_team):
        sum_line += len(globals()['Member' + str(i)])
    for i in range(0, len(TitleItem)):
        SheetOne.write(0, i, TitleItem[i], format)
    merge_format = WorkBook.add_format({'align': 'center', 'valign': 'vcenter'})
    merge_format.set_border(1)
    i, j = 1, 0
    while i < sum_line and j < len(TeamLeader):
        SheetOne.merge_range(i, 0, i - 1 + len(globals()['Member' + str(j)]), 0, TeamLeader[j], merge_format)
        i += len(globals()['Member' + str(j)])
        j += 1
    i = 1
    while i < sum_line:
        for j in range(0, length_team):
            for k in range(0, len(globals()['Member' + str(j)])):
                SheetOne.write(i, 1, (globals()['Member' + str(j)])[k], format)
                ListUsername.append((globals()['Member' + str(j)])[k])
                i += 1
    data_display = {}
    for name in ListUsername:
        data_display['%s' % name] = {}
        data_display['%s' % name]['发明提交数量'.decode('gbk')] = 0
        data_display['%s' % name]['发明受理数量'.decode('gbk')] = 0
        data_display['%s' % name]['实用新型提交数量'.decode('gbk')] = 0
        data_display['%s' % name]['实用新型受理数量'.decode('gbk')] = 0
    # file_name = pandas.read_csv(filename_input, sep=',', header=1, encoding='gbk', na_filter=True, chunksize=1)
    file_name = xlrd.open_workbook(filename_input, encoding_override='cp936')
    sheet_filter_one = file_name.sheet_by_index(0)
    total_rows_one = sheet_filter_one.nrows

    sheet_filter_two = file_name.sheet_by_index(1)
    total_rows_two = sheet_filter_two.nrows

    for item_1 in range(1, total_rows_one):
        department = sheet_filter_one.cell(item_1, 3).value.replace(u' ', u'')
        username = sheet_filter_one.cell(item_1, 6).value.replace(u' ', u'')
        type_invention = sheet_filter_one.cell(item_1, 4).value.replace(u' ', u'')
        shouli_or_not = sheet_filter_one.cell(item_1, 8).value
        if department == department_to_filter:
            if username in ListUsername:
                if shouli_or_not != 'None':
                    if type_invention == '发明'.decode('gbk'):
                        data_display['%s' % username]['发明受理数量'.decode('gbk')] += 1
                    if type_invention == '新型'.decode('gbk'):
                        data_display['%s' % username]['实用新型受理数量'.decode('gbk')] += 1
                if type_invention == '发明'.decode('gbk'):
                    data_display['%s' % username]['发明提交数量'.decode('gbk')] += 1
                if type_invention == '新型'.decode('gbk'):
                    data_display['%s' % username]['实用新型提交数量'.decode('gbk')] += 1

    for item_2 in range(1, total_rows_two):
        department = sheet_filter_two.cell(item_2, 0).value.replace(u' ', u'')
        username = sheet_filter_two.cell(item_2, 1).value
        type_invention = sheet_filter_two.cell(item_2, 6).value.replace(u' ', u'')
        if department == department_to_filter:
            if username in ListUsername:
                if type_invention == '发明'.decode('gbk'):
                    data_display['%s' % username]['发明受理数量'.decode('gbk')] += 1
                if type_invention == '实用新型'.decode('gbk'):
                    data_display['%s' % username]['实用新型受理数量'.decode('gbk')] += 1
    SheetOne.set_column("C:F", 15)
    i = 1
    for username in ListUsername:
        SheetOne.write(i, 2, data_display['%s' % username]['发明受理数量'.decode('gbk')])
        SheetOne.write(i, 3, data_display['%s' % username]['发明提交数量'.decode('gbk')])
        SheetOne.write(i, 4, data_display['%s' % username]['实用新型受理数量'.decode('gbk')])
        SheetOne.write(i, 5, data_display['%s' % username]['实用新型提交数量'.decode('gbk')])
        i += 1
    WorkBook.close()
    tkMessageBox.showinfo('提示'.decode('gbk'), '处理%s的结果已经生成，请去%s路径查看.如果无需其他动作，请点击退出按钮退出程序'.decode('gbk') % (
    department_to_filter, dir_filename_display))


Tkinter.Label(root, text='请在如下选择需要处理的处名'.decode('gbk'), bg='Red').grid(row=0, column=0, columnspan=20, padx=5, pady=5)
box_set_department = ttk.Combobox(root, textvariable=var_char_combox_department,
                                  values=['浪潮集团浪潮信息测试验证部测试二处'.decode('gbk')])
box_set_department.grid(row=1, column=0, columnspan=40, padx=5, pady=5)

Tkinter.Label(root, text='请在如下选择需要处理的专利文件'.decode('gbk'), bg='Red').grid(row=2, column=0, columnspan=20, padx=5, pady=5)
Tkinter.Entry(root, textvariable=var_char_entry_filename_need_filter).grid(row=3, column=0, columnspan=16, padx=10,
                                                                           pady=5)
Tkinter.Button(root, text='选择文件'.decode('gbk'), command=get_filename).grid(row=3, column=16, columnspan=4, padx=5,
                                                                           pady=5, sticky='wesn')

Tkinter.Label(root, text='请在如下选择处理结果存放的位置'.decode('gbk'), bg='Red').grid(row=4, column=0, columnspan=20, padx=5, pady=5)
Tkinter.Entry(root, textvariable=var_char_entry_filename_after_filter).grid(row=5, column=0, columnspan=16, padx=10,
                                                                            pady=5)
Tkinter.Button(root, text='选择文件'.decode('gbk'), command=set_filename).grid(row=5, column=16, columnspan=4, padx=5,
                                                                           pady=5, sticky='wesn')

Tkinter.Button(root, text='GO'.decode('gbk'), command=get_data).grid(row=6, column=0, columnspan=9, padx=10, pady=5,
                                                                     sticky='wesn')
Tkinter.Button(root, text='退出'.decode('gbk'), command=root.destroy).grid(row=6, column=10, columnspan=9, padx=10,
                                                                         pady=5, sticky='wesn')
Tkinter.mainloop()
