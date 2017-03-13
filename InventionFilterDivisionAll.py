#!/bin/usr/env python
# -*- coding:cp936 -*-
"""������ߵ������ǹ��˲��ŵ�ר��������Ȼ��ͳ�Ƴ�g�������˵�ר������������Ϊ���У������ύ����������ʵ�������ύ��ʵ���������������ĵ�
������csv��ʽ�ģ���ֻ��һ��sheet������ļ���Ϊ��������֤�����Ը���ר��������ͳ��.csv��,���ݲ�ͬ�Ĵ������ͬ�ڲ�ͬ��sheet�� �����߻���Python Tkinter����ͼ�ν��档�������import���֡�
�����exe��ʽ��ʹ��pyinstall,����ΪPython pyinstaller.py -F InvertionFilter.py
"""
import pandas
import Tkinter
import tkMessageBox
import  tkFileDialog
import xlsxwriter
import time


filename_original = unicode()
#listDivisionName=['����һ��'.decode('gbk'), '���Զ���'.decode('gbk'), '��������'.decode('gbk'), '�����Ĵ�'.decode('gbk'), '�����崦'.decode('gbk'), '��������'.decode('gbk')]
listDivisionName=['����һ��'.decode('gbk'), '���Զ���'.decode('gbk'), '��������'.decode('gbk')]
listTitle = ['�����ύ����'.decode('gbk'), '������������'.decode('gbk'), 'ʵ�������ύ����'.decode('gbk'), 'ʵ��������������'.decode('gbk')]
root = Tkinter.Tk()
root.title("ר��������˹���".decode('gbk'))
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
    WorkBook = xlsxwriter.Workbook("������֤������ר��������ͳ��.xlsx".decode('gbk'))
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
            data_display['%s' % name]['�����ύ����'.decode('gbk')] = 0
            data_display['%s' % name]['������������'.decode('gbk')] = 0
            data_display['%s' % name]['ʵ�������ύ����'.decode('gbk')] = 0
            data_display['%s' % name]['ʵ��������������'.decode('gbk')] = 0
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
                        if type_invention == '����'.decode('gbk'):
                            data_display['%s' % username]['������������'.decode('gbk')] += 1
                        if type_invention == 'ʵ������'.decode('gbk'):
                            data_display['%s' % username]['ʵ��������������'.decode('gbk')] += 1
                    if backup_1 not in ['����'.decode('gbk'), '15'.decode('gbk')]:
                        if type_invention == '����'.decode('gbk'):
                            data_display['%s' % username]['�����ύ����'.decode('gbk')] += 1
                        if type_invention == 'ʵ������'.decode('gbk'):
                            data_display['%s' % username]['ʵ�������ύ����'.decode('gbk')] += 1
#        print data_display
        for index, item in enumerate(list_username):
            line_count =  index+1
            sheet_now.write(line_count, 0, item, formatOne)
            sheet_now.write(line_count, 1, data_display['%s' % item]['�����ύ����'.decode('gbk')], formatOne)
            sheet_now.write(line_count, 2, data_display['%s' % item]['������������'.decode('gbk')], formatOne)
            sheet_now.write(line_count, 3, data_display['%s' % item]['ʵ�������ύ����'.decode('gbk')], formatOne)
            sheet_now.write(line_count, 4, data_display['%s' % item]['ʵ��������������'.decode('gbk')], formatOne)

    print time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
    tkMessageBox.showinfo('��ʾ'.decode('gbk'), '�������Ѿ���ǰ�ļ���������.����������������������˳���ť�˳�����'.decode('gbk'))
    WorkBook.close()


frame_middle = Tkinter.Frame(root, height=20)
frame_middle.pack(side=Tkinter.TOP)
frame_middle_top = Tkinter.Frame(frame_middle, height=40)
frame_middle_top.pack()
frame_middle_bottom = Tkinter.Frame(frame_middle, height=20)
frame_middle_bottom.pack()
Tkinter.Label(frame_middle_top, text='��������ѡ����Ҫ�����ר���ļ�'.decode('gbk'), bg='Red').pack()

Tkinter.Entry(frame_middle_bottom, textvariable=var_char_entry_filename_need_filter, width=40).pack(side=Tkinter.LEFT)
Tkinter.Button(frame_middle_bottom, text='ѡ���ļ�'.decode('gbk'), command=get_filename, width=20).pack(side=Tkinter.RIGHT)


frame_bottom = Tkinter.Frame(root, height=20)
frame_bottom.pack()

Tkinter.Button(frame_bottom, text='GO'.decode('gbk'), command=get_data, width=20,).pack(side=Tkinter.LEFT)
Tkinter.Button(frame_bottom, text='�˳�'.decode('gbk'), width=20, command=root.destroy).pack(side=Tkinter.LEFT)

Tkinter.mainloop()
