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


filename_original = unicode()
#listDivisionName=['测试一处'.decode('gbk'), '测试二处'.decode('gbk'), '测试三处'.decode('gbk'), '测试四处'.decode('gbk'), '测试五处'.decode('gbk'), '测试六处'.decode('gbk')]
listDivisionName=['测试一处'.decode('gbk'), '测试二处'.decode('gbk'), '测试三处'.decode('gbk')]
listTitle = ['发明提交数量'.decode('gbk'), '发明受理数量'.decode('gbk'), '实用新型提交数量'.decode('gbk'), '实用新型受理数量'.decode('gbk')]
root = Tkinter.Tk()
root.title("专利结果过滤工具".decode('gbk'))
root.geometry('800x600')
root.resizable(width=True, height=True)
var_char_entry_filename_need_filter = Tkinter.StringVar()


def get_filename():
    global filename_original
    filename_iometer = tkFileDialog.askopenfilename()
    var_char_entry_filename_need_filter.set(filename_iometer)
    filename_original = filename_iometer


def get_data():
    print time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
    filename_input = filename_original
    WorkBook = xlsxwriter.Workbook("测试验证部个人专利完成情况统计.xlsx".decode('gbk'))
    formatOne = WorkBook.add_format()
    formatOne.set_border(1)
    for item in listDivisionName:
        WorkBook.add_worksheet(item)
    for department_to_filter in listDivisionName:
        sheet_now = WorkBook.get_worksheet_by_name(department_to_filter)
        for i in range(1, len(listTitle)+1):
            sheet_now.write(0, i, listTitle[i-1], formatOne)
        list_username = []
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
#        print data_display
        for index, item in enumerate(list_username):
            line_count =  index+1
            sheet_now.write(line_count, 0, item, formatOne)
            sheet_now.write(line_count, 1, data_display['%s' % item]['发明提交数量'.decode('gbk')], formatOne)
            sheet_now.write(line_count, 2, data_display['%s' % item]['发明受理数量'.decode('gbk')], formatOne)
            sheet_now.write(line_count, 3, data_display['%s' % item]['实用新型提交数量'.decode('gbk')], formatOne)
            sheet_now.write(line_count, 4, data_display['%s' % item]['实用新型受理数量'.decode('gbk')], formatOne)

    print time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
    tkMessageBox.showinfo('提示'.decode('gbk'), '处理结果已经当前文件夹下生成.如果无需其他动作，请点击退出按钮退出程序'.decode('gbk'))
    WorkBook.close()


frame_middle = Tkinter.Frame(root, height=20)
frame_middle.pack(side=Tkinter.TOP)
frame_middle_top = Tkinter.Frame(frame_middle, height=40)
frame_middle_top.pack()
frame_middle_bottom = Tkinter.Frame(frame_middle, height=20)
frame_middle_bottom.pack()
Tkinter.Label(frame_middle_top, text='请在如下选择需要处理的专利文件'.decode('gbk'), bg='Red').pack()

Tkinter.Entry(frame_middle_bottom, textvariable=var_char_entry_filename_need_filter, width=40).pack(side=Tkinter.LEFT)
Tkinter.Button(frame_middle_bottom, text='选择文件'.decode('gbk'), command=get_filename, width=20).pack(side=Tkinter.RIGHT)


frame_bottom = Tkinter.Frame(root, height=20)
frame_bottom.pack()

Tkinter.Button(frame_bottom, text='GO'.decode('gbk'), command=get_data, width=20,).pack(side=Tkinter.LEFT)
Tkinter.Button(frame_bottom, text='退出'.decode('gbk'), width=20, command=root.destroy).pack(side=Tkinter.LEFT)

Tkinter.mainloop()
