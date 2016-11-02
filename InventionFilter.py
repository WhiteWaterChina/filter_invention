#!/bin/usr/env python
# -*- coding:cp936 -*-
"""这个工具的作用是过滤部门的专利总览表，然后统计出g各处个人的专利完成情况。分为四列：发明提交、发明受理、实用新型提交、实用新型受理。输入文档
必须是csv格式的，且只有一个sheet。输出文件名为《测试验证部测试%s个人专利完成情况统计.csv》,依据不同的处别而不同。 本工具基于Python Tkinter制作图形界面。依赖详见import部分。
打包成exe格式请使用pyinstall,命令为Python pyinstaller.py -F InvertionFilter.py
"""
import pandas
import Tkinter
import tkMessageBox
import ttk
import  tkFileDialog
import os


filename_original = unicode()
dir_filename_display = unicode()
root = Tkinter.Tk()
root.title("专利结果过滤工具".decode('gbk'))
root.geometry('800x600')
root.resizable(width=True, height=True)
var_char_entry_filename_need_filter = Tkinter.StringVar()
var_char_combox_department = Tkinter.StringVar()
var_char_entry_filename_after_filter = Tkinter.StringVar()


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
    list_username = []
    filename_input = filename_original
    department_to_filter = var_char_combox_department.get()
    filename_output = os.path.join(dir_filename_display, "测试验证部测试%s个人专利完成情况统计.csv".decode('gbk') % department_to_filter)
    file_name = pandas.read_csv(filename_input, sep=',', header=1, encoding='gbk', na_filter=True, chunksize=1)
    for item in file_name:
        department = item.iloc[0, 0].replace(u' ', u'')
        if department == department_to_filter:
            username = item.iloc[0, 1]
            if username not in list_username:
                list_username.append(username)
    data_display = {}
    for name in list_username:
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
            if username in list_username:
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
    dataframe_data = pandas.DataFrame(data_display).T
    dataframe_data.to_csv(filename_output, encoding='gbk')
    tkMessageBox.showinfo('提示'.decode('gbk'), '处理%s的结果已经生成，请去%s路径查看.如果无需其他动作，请点击退出按钮退出程序'.decode('gbk') % (department_to_filter, dir_filename_display))

frame_top = Tkinter.Frame(root, height=20)
frame_top.pack(side=Tkinter.TOP)

Tkinter.Label(frame_top, text='请在如下选择需要处理的处名'.decode('gbk'), bg='Red').pack(side=Tkinter.TOP)
box_set_department = ttk.Combobox(frame_top, textvariable=var_char_combox_department,
                                  values=['测试一处'.decode('gbk'), '测试二处'.decode('gbk'), '测试三处'.decode('gbk'),
                                          '测试四处'.decode('gbk'), '测试五处'.decode('gbk'), '测试六处'.decode('gbk')], width=30)

box_set_department.pack(side=Tkinter.BOTTOM)

frame_middle = Tkinter.Frame(root, height=20)
frame_middle.pack()
frame_middle_top = Tkinter.Frame(frame_middle, height=40)
frame_middle_top.pack()
frame_middle_bottom = Tkinter.Frame(frame_middle, height=20)
frame_middle_bottom.pack()
Tkinter.Label(frame_middle_top, text='请在如下选择需要处理的专利文件'.decode('gbk'), bg='Red').pack()

Tkinter.Entry(frame_middle_bottom, textvariable=var_char_entry_filename_need_filter, width=40).pack(side=Tkinter.LEFT)
Tkinter.Button(frame_middle_bottom, text='选择文件'.decode('gbk'), command=get_filename, width=20).pack(side=Tkinter.RIGHT)

frame_middle_2 = Tkinter.Frame(root, height=20)
frame_middle_2.pack()
frame_middle_top_2 = Tkinter.Frame(frame_middle_2, height=40)
frame_middle_top_2.pack()
frame_middle_bottom_2 = Tkinter.Frame(frame_middle_2, height=20)
frame_middle_bottom_2.pack()
Tkinter.Label(frame_middle_top_2, text='请在如下选择处理结果存放的位置'.decode('gbk'), bg='Red').pack()
Tkinter.Entry(frame_middle_bottom_2, textvariable=var_char_entry_filename_after_filter, width=40).pack(side=Tkinter.LEFT)
Tkinter.Button(frame_middle_bottom_2, text='选择文件'.decode('gbk'), command=set_filename, width=20).pack(side=Tkinter.RIGHT)


frame_bottom = Tkinter.Frame(root, height=20)
frame_bottom.pack()

Tkinter.Button(frame_bottom, text='GO'.decode('gbk'), width=20, command=get_data).pack(side=Tkinter.LEFT)
Tkinter.Button(frame_bottom, text='退出'.decode('gbk'), width=20, command=root.destroy).pack(side=Tkinter.LEFT)

Tkinter.mainloop()
