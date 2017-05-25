#!/bin/usr/env python
# -*- coding:cp936 -*-
"""������ߵ������ǹ��˲��ŵ�ר��������Ȼ��ͳ�Ƴ�һ�Ͷ������˵�ר������������Ϊ���У��鳤������Ա���������ύ����������ʵ�������ύ��ʵ���������������ĵ�
������csv��ʽ�ģ���ֻ��һ��sheet������ļ���Ϊ��������֤������%s����ר��������ͳ��.csv��,���ݲ�ͬ�Ĵ������ͬ�� �����߻���Python Tkinter����ͼ�ν��档�������import���֡�
�����exe��ʽ��ʹ��pyinstall,����ΪPython pyinstaller.py -F InventionFilterDepartmentOne.py
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
root.title("ר��������˹���".decode('gbk'))
# root.geometry('800x600')
root.resizable(width=True, height=True)
var_char_entry_filename_need_filter = Tkinter.StringVar()
var_char_combox_department = Tkinter.StringVar()
var_char_entry_filename_after_filter = Tkinter.StringVar()

TeamLeader = ['�ֵ�'.decode('gbk'), '����'.decode('gbk'), '������'.decode('gbk'), '������'.decode('gbk'),
              'ʷ����'.decode('gbk'), '������'.decode('gbk'), '������'.decode('gbk'), '�ٽ���'.decode('gbk'), '������'.decode('gbk'),
              '����'.decode('gbk')]
# �����ְ
Member0 = ['�ֵ�'.decode('gbk'), '����'.decode('gbk'), '��ï��'.decode('gbk'), '������'.decode('gbk'), '̷����'.decode('gbk'),
           '������'.decode('gbk'), '���羲'.decode('gbk')]
Member1 = ['����'.decode('gbk'), '����'.decode('gbk'), '����'.decode('gbk'), '��ϲ��'.decode('gbk'), '������'.decode('gbk')]
Member2 = ['������'.decode('gbk'), '�ֺ�'.decode('gbk'), '�ܼӷ�'.decode('gbk'), '������'.decode('gbk'), '���'.decode('gbk'),
           '������'.decode('gbk')]
Member3 = ['������'.decode('gbk'), '��˶'.decode('gbk'), '����ѩ'.decode('gbk'), '������'.decode('gbk'), '¬����'.decode('gbk')]
Member4 = ['ʷ����'.decode('gbk'), '�ų�'.decode('gbk'), '���'.decode('gbk'), '���Ǹ�'.decode('gbk'), '������'.decode('gbk'),
           '����'.decode('gbk'), '����'.decode('gbk'), '��ΰ��'.decode('gbk'), '��ʢ'.decode('gbk'), '������'.decode('gbk'),
           '��Ө'.decode('gbk'), '������'.decode('gbk'), '���'.decode('gbk'), '�̼Ѽ�'.decode('gbk'),]
Member5 = ['������'.decode('gbk'), '�O�۱�'.decode('gbk'), '��ޱ'.decode('gbk'), '�'.decode('gbk'), '������'.decode('gbk')]
Member6 = ['������'.decode('gbk'), '�'.decode('gbk'), '����ΰ'.decode('gbk'), '������'.decode('gbk'), '������'.decode('gbk'),
           '��־Ȫ'.decode('gbk'), '����'.decode('gbk'), '��׳'.decode('gbk'), '���'.decode('gbk')]
Member7 = ['�ٽ���'.decode('gbk'), '���ƾ�'.decode('gbk'), '���廪'.decode('gbk'), '������'.decode('gbk'),
            '����ѫ'.decode('gbk'), '�ƽ�ʢ'.decode('gbk')]
Member8 = ['������'.decode('gbk'), '�6011'.decode('gbk'), '��̫˳'.decode('gbk')]
Member9 = ['����'.decode('gbk'), '������'.decode('gbk'), '������'.decode('gbk'), '�����'.decode('gbk'), '�Y����'.decode('gbk')]

TitleItem = ['�鳤��'.decode('gbk'), '��Ա��'.decode('gbk'), '������������'.decode('gbk'), '�����ύ����'.decode('gbk'),
             'ʵ��������������'.decode('gbk'), 'ʵ�������ύ����'.decode('gbk')]


def get_filename():
    global filename_original
    filename_iometer = tkFileDialog.askopenfilename()
    var_char_entry_filename_need_filter.set(filename_iometer)
    filename_original = filename_iometer


def set_filename():
    global dir_filename_display
    dir_filename_display = tkFileDialog.askdirectory().replace('/', '\\')
    var_char_entry_filename_after_filter.set(dir_filename_display)
#    filename_display = os.path.join(dir_filter_iometer, "������֤������%s����ר��������ͳ��.csv".decode('gbk') )


def get_data():
    length_team = len(TeamLeader)
    filename_input = filename_original
    department_to_filter = var_char_combox_department.get()
    timestamp = time.strftime('%Y%m%d', time.localtime())
    filename_output = os.path.join(dir_filename_display,
                                   "%s����ר��������ͳ��-%s.xlsx".decode('gbk') % (department_to_filter, timestamp))
    WorkBook = xlsxwriter.Workbook(filename_output)
    SheetOne = WorkBook.add_worksheet('����ר��������ͳ��'.decode('gbk'))
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
        data_display['%s' % name]['�����ύ����'.decode('gbk')] = 0
        data_display['%s' % name]['������������'.decode('gbk')] = 0
        data_display['%s' % name]['ʵ�������ύ����'.decode('gbk')] = 0
        data_display['%s' % name]['ʵ��������������'.decode('gbk')] = 0
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
                    if type_invention == '����'.decode('gbk'):
                        data_display['%s' % username]['������������'.decode('gbk')] += 1
                    if type_invention == '����'.decode('gbk'):
                        data_display['%s' % username]['ʵ��������������'.decode('gbk')] += 1
                if type_invention == '����'.decode('gbk'):
                    data_display['%s' % username]['�����ύ����'.decode('gbk')] += 1
                if type_invention == '����'.decode('gbk'):
                    data_display['%s' % username]['ʵ�������ύ����'.decode('gbk')] += 1

    for item_2 in range(1, total_rows_two):
        department = sheet_filter_two.cell(item_2, 0).value.replace(u' ', u'')
        username = sheet_filter_two.cell(item_2, 1).value
        type_invention = sheet_filter_two.cell(item_2, 6).value.replace(u' ', u'')
        if department == department_to_filter:
            if username in ListUsername:
                if type_invention == '����'.decode('gbk'):
                    data_display['%s' % username]['������������'.decode('gbk')] += 1
                if type_invention == 'ʵ������'.decode('gbk'):
                    data_display['%s' % username]['ʵ��������������'.decode('gbk')] += 1
    SheetOne.set_column("C:F", 15)
    i = 1
    for username in ListUsername:
        SheetOne.write(i, 2, data_display['%s' % username]['������������'.decode('gbk')])
        SheetOne.write(i, 3, data_display['%s' % username]['�����ύ����'.decode('gbk')])
        SheetOne.write(i, 4, data_display['%s' % username]['ʵ��������������'.decode('gbk')])
        SheetOne.write(i, 5, data_display['%s' % username]['ʵ�������ύ����'.decode('gbk')])
        i += 1
    WorkBook.close()
    tkMessageBox.showinfo('��ʾ'.decode('gbk'), '����%s�Ľ���Ѿ����ɣ���ȥ%s·���鿴.����������������������˳���ť�˳�����'.decode('gbk') % (
    department_to_filter, dir_filename_display))


Tkinter.Label(root, text='��������ѡ����Ҫ����Ĵ���'.decode('gbk'), bg='Red').grid(row=0, column=0, columnspan=20, padx=5, pady=5)
box_set_department = ttk.Combobox(root, textvariable=var_char_combox_department,
                                  values=['�˳������˳���Ϣ������֤�����Զ���'.decode('gbk')])
box_set_department.grid(row=1, column=0, columnspan=40, padx=5, pady=5)

Tkinter.Label(root, text='��������ѡ����Ҫ�����ר���ļ�'.decode('gbk'), bg='Red').grid(row=2, column=0, columnspan=20, padx=5, pady=5)
Tkinter.Entry(root, textvariable=var_char_entry_filename_need_filter).grid(row=3, column=0, columnspan=16, padx=10,
                                                                           pady=5)
Tkinter.Button(root, text='ѡ���ļ�'.decode('gbk'), command=get_filename).grid(row=3, column=16, columnspan=4, padx=5,
                                                                           pady=5, sticky='wesn')

Tkinter.Label(root, text='��������ѡ��������ŵ�λ��'.decode('gbk'), bg='Red').grid(row=4, column=0, columnspan=20, padx=5, pady=5)
Tkinter.Entry(root, textvariable=var_char_entry_filename_after_filter).grid(row=5, column=0, columnspan=16, padx=10,
                                                                            pady=5)
Tkinter.Button(root, text='ѡ���ļ�'.decode('gbk'), command=set_filename).grid(row=5, column=16, columnspan=4, padx=5,
                                                                           pady=5, sticky='wesn')

Tkinter.Button(root, text='GO'.decode('gbk'), command=get_data).grid(row=6, column=0, columnspan=9, padx=10, pady=5,
                                                                     sticky='wesn')
Tkinter.Button(root, text='�˳�'.decode('gbk'), command=root.destroy).grid(row=6, column=10, columnspan=9, padx=10,
                                                                         pady=5, sticky='wesn')
Tkinter.mainloop()
