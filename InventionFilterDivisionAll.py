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
import os


filename_original = unicode()
#listDivisionName=['����һ��'.decode('gbk'), '���Զ���'.decode('gbk'), '��������'.decode('gbk'), '�����Ĵ�'.decode('gbk'), '�����崦'.decode('gbk'), '��������'.decode('gbk')]
listDivisionName=['����һ��'.decode('gbk'), '���Զ���'.decode('gbk'), '��������'.decode('gbk'), '�����Ĵ�'.decode('gbk'), '�����崦'.decode('gbk'), '��������'.decode('gbk')]
listTitle = ['�����ύ����'.decode('gbk'), '������������'.decode('gbk'), 'ʵ�������ύ����'.decode('gbk'), 'ʵ��������������'.decode('gbk')]
root = Tkinter.Tk()
root.title("ר��������˹���-���д�".decode('gbk'))
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
    filename_final = os.path.join(dir_filename_display, "������֤������ר��������ͳ��-%s.xlsx".decode('gbk') % timestamp)
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
            data_display['%s' % name]['�����ύ����'.decode('gbk')] = 0
            data_display['%s' % name]['������������'.decode('gbk')] = 0
            data_display['%s' % name]['ʵ�������ύ����'.decode('gbk')] = 0
            data_display['%s' % name]['ʵ��������������'.decode('gbk')] = 0
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
                        if type_invention == '����'.decode('gbk'):
                            data_display['%s' % username]['������������'.decode('gbk')] += 1
                        if type_invention == '����'.decode('gbk'):
                            data_display['%s' % username]['ʵ��������������'.decode('gbk')] += 1
#                    if backup_1 not in ['����'.decode('gbk'), '15'.decode('gbk')]:
                    if type_invention == '����'.decode('gbk'):
                        data_display['%s' % username]['�����ύ����'.decode('gbk')] += 1
                    if type_invention == '����'.decode('gbk'):
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
    tkMessageBox.showinfo('��ʾ'.decode('gbk'), '�������Ѿ����浽�ļ�%s��.����������������������˳���ť�˳�����'.decode('gbk') % filename_final)
    WorkBook.close()

Tkinter.Label(root, text='��������ѡ����Ҫ�����ר���ļ�'.decode('gbk'), bg='Red').grid(row=0, column=0, columnspan=5, padx=10, pady=5)
Tkinter.Entry(root, textvariable=var_char_entry_filename_need_filter).grid(row=1, column=0, columnspan=4, padx=5, pady=5, sticky='we')
Tkinter.Button(root, text='ѡ���ļ�'.decode('gbk'), command=get_filename).grid(row=1, column=4, padx=5, pady=5)

Tkinter.Label(root, text='��������ѡ����������Ŀ¼'.decode('gbk'), bg='Red').grid(row=2, column=0, columnspan=5, padx=10, pady=5)
Tkinter.Entry(root, textvariable=var_char_entry_filename_after_filter).grid(row=3, column=0, columnspan=4, padx=5, pady=5, sticky='we')
Tkinter.Button(root, text='ѡ��Ŀ¼'.decode('gbk'), command=set_filename).grid(row=3, column=4, padx=5, pady=5)

Tkinter.Button(root, text='GO'.decode('gbk'), command=get_data).grid(row=4, column=0, columnspan=2, padx=5, pady=5, sticky='wesn')
Tkinter.Button(root, text='�˳�'.decode('gbk'), command=root.destroy).grid(row=4, column=3, columnspan=2, padx=5, pady=5, sticky='wesn')

Tkinter.mainloop()
