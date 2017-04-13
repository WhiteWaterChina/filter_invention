#!/bin/usr/env python
# -*- coding:cp936 -*-
"""这个工具的作用是过滤部门的专利总览表，然后统计出g各处个人的专利完成情况。分为四列：发明提交、发明受理、实用新型提交、实用新型受理。输入文档
必须是csv格式的，且只有一个sheet。输出文件名为《测试验证部测试个人专利完成情况统计.csv》,依据不同的处别而不同在不同的sheet。 本工具基于Python Tkinter制作图形界面。依赖详见import部分。
打包成exe格式请使用pyinstall,命令为Python pyinstaller.py -F InvertionFilter.py
"""
import pandas
import Tkinter
import tkMessageBox
import  tkFileDialog
import xlsxwriter
import time
import os


filename_original = unicode()
#listDivisionName=['测试一处'.decode('gbk'), '测试二处'.decode('gbk'), '测试三处'.decode('gbk'), '测试四处'.decode('gbk'), '测试五处'.decode('gbk'), '测试六处'.decode('gbk')]
listDivisionName=['测试一处'.decode('gbk'), '测试二处'.decode('gbk'), '测试三处'.decode('gbk'), '测试四处'.decode('gbk'), '测试五处'.decode('gbk'), '测试六处'.decode('gbk')]
listTitle = ['发明提交数量'.decode('gbk'), '发明受理数量'.decode('gbk'), '实用新型提交数量'.decode('gbk'), '实用新型受理数量'.decode('gbk')]
root = Tkinter.Tk()
root.title("专利结果过滤工具-所有处".decode('gbk'))
#root.geometry('400x300')
#root.iconbitmap('logo.ico')
root.resizable(width=True, height=True)
var_char_entry_filename_need_filter = Tkinter.StringVar()
var_char_entry_filename_after_filter = Tkinter.StringVar()


def get_filename():
    global filename_original
    filename_invention = tkFileDialog.askopenfilename()
    var_char_entry_filename_need_filter.set(filename_invention)
    filename_original = filename_invention


def set_filename():
    global dir_filename_display
    dir_filename_display = tkFileDialog.askdirectory().replace('/', '\\')
    var_char_entry_filename_after_filter.set(dir_filename_display)


def get_data():
    print time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
    filename_input = filename_original
    timestamp = time.strftime('%Y%m%d', time.localtime())
    filename_final = os.path.join(dir_filename_display, "测试验证部个人专利完成情况统计-%s.xlsx".decode('gbk') % timestamp)
    WorkBook = xlsxwriter.Workbook(filename_final)
    formatOne = WorkBook.add_format()
    formatOne.set_border(1)
    for item in listDivisionName:
        WorkBook.add_worksheet(item)
    for department_to_filter in listDivisionName:
        sheet_now = WorkBook.get_worksheet_by_name(department_to_filter)
        sheet_now.set_column('B:E', 15)
        for i in range(1, len(listTitle)+1):
            sheet_now.write(0, i, listTitle[i-1], formatOne)
        list_username = []
        file_name = pandas.read_csv(filename_input, sep=',', header=1, encoding='gbk', na_filter=True, chunksize=1)
        for item in file_name:
            department = item.iloc[0, 3].replace(u' ', u'')
            if department == department_to_filter:
                username = item.iloc[0, 6]
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
            department = item_1.iloc[0, 3].replace(u' ', u'')
            username = item_1.iloc[0, 6].replace(u' ', u'')
            type_invention = item_1.iloc[0, 4].replace(u' ', u'')
            date_shouli = item_1.iloc[0, 8]
#            backup = item_1.iloc[0, 8]
#            backup_1 = 0
#            if pandas.notnull(backup):
#                backup_1 = backup[0:2]

            if department == department_to_filter:
                if username in list_username:
                    if pandas.notnull(date_shouli):
                        if type_invention == '发明'.decode('gbk'):
                            data_display['%s' % username]['发明受理数量'.decode('gbk')] += 1
                        if type_invention == '新型'.decode('gbk'):
                            data_display['%s' % username]['实用新型受理数量'.decode('gbk')] += 1
#                    if backup_1 not in ['放弃'.decode('gbk'), '15'.decode('gbk')]:
                    if type_invention == '发明'.decode('gbk'):
                        data_display['%s' % username]['发明提交数量'.decode('gbk')] += 1
                    if type_invention == '新型'.decode('gbk'):
                        data_display['%s' % username]['实用新型提交数量'.decode('gbk')] += 1
#        print data_display
        for index, item in enumerate(list_username):
            line_count =  index+1
            sheet_now.write(line_count, 0, item, formatOne)
            sheet_now.write(line_count, 1, data_display['%s' % item]['发明提交数量'.decode('gbk')], formatOne)
            sheet_now.write(line_count, 2, data_display['%s' % item]['发明受理数量'.decode('gbk')], formatOne)
            sheet_now.write(line_count, 3, data_display['%s' % item]['实用新型提交数量'.decode('gbk')], formatOne)
            sheet_now.write(line_count, 4, data_display['%s' % item]['实用新型受理数量'.decode('gbk')], formatOne)

    print time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
    tkMessageBox.showinfo('提示'.decode('gbk'), '处理结果已经保存到文件%s中.如果无需其他动作，请点击退出按钮退出程序'.decode('gbk') % filename_final)
    WorkBook.close()

Tkinter.Label(root, text='请在如下选择需要处理的专利文件'.decode('gbk'), bg='Red').grid(row=0, column=0, columnspan=5, padx=10, pady=5)
Tkinter.Entry(root, textvariable=var_char_entry_filename_need_filter).grid(row=1, column=0, columnspan=4, padx=5, pady=5, sticky='we')
Tkinter.Button(root, text='选择文件'.decode('gbk'), command=get_filename).grid(row=1, column=4, padx=5, pady=5)

Tkinter.Label(root, text='请在如下选择输出结果的目录'.decode('gbk'), bg='Red').grid(row=2, column=0, columnspan=5, padx=10, pady=5)
Tkinter.Entry(root, textvariable=var_char_entry_filename_after_filter).grid(row=3, column=0, columnspan=4, padx=5, pady=5, sticky='we')
Tkinter.Button(root, text='选择目录'.decode('gbk'), command=set_filename).grid(row=3, column=4, padx=5, pady=5)

Tkinter.Button(root, text='GO'.decode('gbk'), command=get_data).grid(row=4, column=0, columnspan=2, padx=5, pady=5, sticky='wesn')
Tkinter.Button(root, text='退出'.decode('gbk'), command=root.destroy).grid(row=4, column=3, columnspan=2, padx=5, pady=5, sticky='wesn')

Tkinter.mainloop()
