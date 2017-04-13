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
import  tkFileDialog
import os
import time
import xlsxwriter


filename_original = unicode()
dir_filename_display = unicode()
root = Tkinter.Tk()
root.title("专利结果过滤工具".decode('gbk'))
#root.geometry('800x600')
root.resizable(width=True, height=True)
var_char_entry_filename_need_filter = Tkinter.StringVar()
var_char_combox_department = Tkinter.StringVar()
var_char_entry_filename_after_filter = Tkinter.StringVar()

TeamLeader = ['贾岛'.decode('gbk'), '刘云飞'.decode('gbk'), '韩琳琳'.decode('gbk'), '苗永威'.decode('gbk'),
              '周志超'.decode('gbk'), '史沛玉'.decode('gbk'), '杨文清'.decode('gbk'), '牟茜'.decode('gbk'), '齐煜'.decode('gbk'),
              '路明远'.decode('gbk'), '王超'.decode('gbk'), '伯绍文'.decode('gbk'), '迟江波'.decode('gbk'),
              '戴明甫'.decode('gbk'), '张晓涛'.decode('gbk'), '曹翔'.decode('gbk')]
# 万浩离职
Member0 = ['贾岛'.decode('gbk'), '李光达'.decode('gbk'), '万浩'.decode('gbk'), '范鹏飞'.decode('gbk'), '谭静静'.decode('gbk'),
           '张文珂'.decode('gbk')]
Member1 = ['刘云飞'.decode('gbk'), '刘博'.decode('gbk'), '潘禹同'.decode('gbk'), '董喜燕'.decode('gbk'), '郝良晟'.decode('gbk')]
Member2 = ['韩琳琳'.decode('gbk'), '林海'.decode('gbk'), '曹加峰'.decode('gbk'), '张行武'.decode('gbk'), '李建波'.decode('gbk'),'冯晓洁'.decode('gbk')]
Member3 = ['苗永威'.decode('gbk'), '闫硕'.decode('gbk'), '刘瑞雪'.decode('gbk'), '王云鹏'.decode('gbk'), '许雪雪'.decode('gbk')]
Member4 = ['周志超'.decode('gbk'), '于兴龙'.decode('gbk'), '谢从波'.decode('gbk'), '姜庆臣'.decode('gbk'), '武琳琳'.decode('gbk')]
# 孟亚男、杨继德、肖欢离职
Member5 = ['史沛玉'.decode('gbk'), '张超'.decode('gbk'), '张锟'.decode('gbk'), '刘智刚'.decode('gbk'), '巩祥文'.decode('gbk'),
           '孙玉超'.decode('gbk'), '韩超'.decode('gbk'), '徐伟超'.decode('gbk'), '赵盛'.decode('gbk'), '肖欢'.decode('gbk'),
           '孟亚男'.decode('gbk'), '杨继德'.decode('gbk'), '王旭林'.decode('gbk'), '杨惠'.decode('gbk'), '程佳佳'.decode('gbk'),
           '刘辉'.decode('gbk')]
# 张娜、黄贤鹤实习生
Member6 = ['杨文清'.decode('gbk'), '潘霖'.decode('gbk'), '孙薇'.decode('gbk'), '李静'.decode('gbk')]
# 于勤伟已经离职
Member7 = ['牟茜'.decode('gbk'), '李萌'.decode('gbk'), '刘振东'.decode('gbk'), '王文悦'.decode('gbk'), '于勤伟'.decode('gbk'),
           '姜敏'.decode('gbk')]
Member8 = ['齐煜'.decode('gbk'), '田立文'.decode('gbk'), '丁凯乐'.decode('gbk')]
Member9 = ['路明远'.decode('gbk'), '王野'.decode('gbk'), '息培磊'.decode('gbk'), '张宇'.decode('gbk'), '伍媚'.decode('gbk')]
Member10 = ['王超'.decode('gbk'), '曲洪磊'.decode('gbk'), '崔夕军'.decode('gbk'), '李强'.decode('gbk'), '戈文龙'.decode('gbk')]
# 于永杰离职
Member11 = ['伯绍文'.decode('gbk'), '李波'.decode('gbk'), '刘东伟'.decode('gbk'), '吴培琴'.decode('gbk'), '于永杰'.decode('gbk'),
            '张贺'.decode('gbk'), '武秋星'.decode('gbk'), '胥志泉'.decode('gbk'), '赵召'.decode('gbk'), '李壮'.decode('gbk'),
            '曹洪帅'.decode('gbk'), '韩刚'.decode('gbk'), '朱敬'.decode('gbk')]
Member12 = ['迟江波'.decode('gbk'), '刘浩君'.decode('gbk'), '周潇'.decode('gbk'), '李彦华'.decode('gbk'), '韩燕燕'.decode('gbk'),
            '梁恒勋'.decode('gbk'), '黄锦盛'.decode('gbk'), '王晓明'.decode('gbk')]
# 吴畏已经离职
Member13 = ['戴明甫'.decode('gbk'), '杨帅'.decode('gbk'), '吴畏'.decode('gbk'), '张常利'.decode('gbk'),
            '李莎莎'.decode('gbk'), '熊婷凤'.decode('gbk')]
# 洪强已经离职
Member14 = ['张晓涛'.decode('gbk'), '洪强'.decode('gbk')]
Member15 = ['曹翔'.decode('gbk'), '李丹'.decode('gbk') , '李永亮'.decode('gbk')]

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
    ListUsername = []
    filename_input = filename_original
    department_to_filter = var_char_combox_department.get()
    timestamp = time.strftime('%Y%m%d', time.localtime())
    filename_output = os.path.join(dir_filename_display, "测试验证部%s个人专利完成情况统计-%s.xlsx".decode('gbk') % (department_to_filter, timestamp))
    WorkBook = xlsxwriter.Workbook(filename_output)
    SheetOne = WorkBook.add_worksheet('测试验证部%s个人专利完成情况统计'.decode('gbk'))
    format = WorkBook.add_format()
    format.set_border(1)
    sum_line = 0
    ListUsername = []
    for i in range(0, 16):
        sum_line += len(globals()['Member'+str(i)])
    for i in range(0, len(TitleItem)):
        SheetOne.write(0, i, TitleItem[i], format)
    merge_format = WorkBook.add_format({'align': 'center', 'valign': 'vcenter'})
    merge_format.set_border(1)
    i, j = 1, 0
    while i < sum_line and j < len(TeamLeader):
        SheetOne.merge_range(i, 0, i - 1 + len(globals()['Member'+str(j)]), 0, TeamLeader[j], merge_format)
        i += len(globals()['Member'+str(j)])
        j += 1
    i = 1
    while i < sum_line:
        for j in range(0, 16):
            for k in range(0, len(globals()['Member'+str(j)])):
                SheetOne.write(i, 1, (globals()['Member'+str(j)])[k], format)
                ListUsername.append((globals()['Member'+str(j)])[k])
                i += 1
    data_display = {}
    for name in ListUsername:
        data_display['%s' % name] = {}
        data_display['%s' % name]['发明提交数量'.decode('gbk')] = 0
        data_display['%s' % name]['发明受理数量'.decode('gbk')] = 0
        data_display['%s' % name]['实用新型提交数量'.decode('gbk')] = 0
        data_display['%s' % name]['实用新型受理数量'.decode('gbk')] = 0
    file_name = pandas.read_csv(filename_input, sep=',', header=1, encoding='gbk', na_filter=True, chunksize=1)
    for item_1 in file_name:
        department = item_1.iloc[0, 0].replace(u' ', u'')
        username = item_1.iloc[0, 1].replace(u' ', u'')
        type_invention = item_1.iloc[0, 6].replace(u' ', u'')
        date_shouli = item_1.iloc[0, 5]
        backup = item_1.iloc[0, 8]
        backup_1 = 0
        if pandas.notnull(backup):
            backup_1 = backup[0:2]
        if department == department_to_filter:
            if username in ListUsername:
                if pandas.notnull(date_shouli):
                    if type_invention == '发明'.decode('gbk'):
                        data_display['%s' % username]['发明受理数量'.decode('gbk')] += 1
                    if type_invention == '实用新型'.decode('gbk'):
                        data_display['%s' % username]['实用新型受理数量'.decode('gbk')] += 1
                if backup_1 not in ['放弃'.decode('gbk'), '15'.decode('gbk')]:
                    if type_invention == '发明'.decode('gbk'):
                        data_display['%s' % username]['发明提交数量'.decode('gbk')] += 1
                    if type_invention == '实用新型'.decode('gbk'):
                        data_display['%s' % username]['实用新型提交数量'.decode('gbk')] += 1
    i = 1
    for username in ListUsername:
        SheetOne.write(i, 2, data_display['%s' % username]['发明受理数量'.decode('gbk')])
        SheetOne.write(i, 3, data_display['%s' % username]['发明提交数量'.decode('gbk')])
        SheetOne.write(i, 4, data_display['%s' % username]['实用新型受理数量'.decode('gbk')])
        SheetOne.write(i, 5, data_display['%s' % username]['实用新型提交数量'.decode('gbk')])
        i += 1
    WorkBook.close()
    tkMessageBox.showinfo('提示'.decode('gbk'), '处理%s的结果已经生成，请去%s路径查看.如果无需其他动作，请点击退出按钮退出程序'.decode('gbk') % (department_to_filter, dir_filename_display))


Tkinter.Label(root, text='请在如下选择需要处理的处名'.decode('gbk'), bg='Red').grid(row=0, column=0, columnspan=20, padx=5, pady=5)
box_set_department = ttk.Combobox(root, textvariable=var_char_combox_department,
                                  values=['测试二处'.decode('gbk')])
box_set_department.grid(row=1, column=0, columnspan=40, padx=5, pady=5)

Tkinter.Label(root, text='请在如下选择需要处理的专利文件'.decode('gbk'), bg='Red').grid(row=2, column=0, columnspan=20, padx=5, pady=5)
Tkinter.Entry(root, textvariable=var_char_entry_filename_need_filter).grid(row=3, column=0, columnspan=16, padx=10, pady=5)
Tkinter.Button(root, text='选择文件'.decode('gbk'), command=get_filename).grid(row=3, column=16, columnspan=4, padx=5, pady=5, sticky='wesn')

Tkinter.Label(root, text='请在如下选择处理结果存放的位置'.decode('gbk'), bg='Red').grid(row=4, column=0, columnspan=20, padx=5, pady=5)
Tkinter.Entry(root, textvariable=var_char_entry_filename_after_filter).grid(row=5, column=0, columnspan=16, padx=10, pady=5)
Tkinter.Button(root, text='选择文件'.decode('gbk'), command=set_filename).grid(row=5, column=16, columnspan=4, padx=5, pady=5, sticky='wesn')

Tkinter.Button(root, text='GO'.decode('gbk'), command=get_data).grid(row=6, column=0, columnspan=9, padx=10, pady=5, sticky='wesn')
Tkinter.Button(root, text='退出'.decode('gbk'), command=root.destroy).grid(row=6, column=10, columnspan=9, padx=10, pady=5, sticky='wesn')
Tkinter.mainloop()
