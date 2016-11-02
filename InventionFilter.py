#!/bin/usr/env python
# -*- coding:cp936 -*-
"""������ߵ������ǹ��˲��ŵ�ר��������Ȼ��ͳ�Ƴ�g�������˵�ר������������Ϊ���У������ύ����������ʵ�������ύ��ʵ���������������ĵ�
������csv��ʽ�ģ���ֻ��һ��sheet������ļ���Ϊ��������֤������%s����ר��������ͳ��.csv��,���ݲ�ͬ�Ĵ������ͬ�� �����߻���Python Tkinter����ͼ�ν��档�������import���֡�
�����exe��ʽ��ʹ��pyinstall,����ΪPython pyinstaller.py -F InvertionFilter.py
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
root.title("ר��������˹���".decode('gbk'))
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
#    filename_display = os.path.join(dir_filter_iometer, "������֤������%s����ר��������ͳ��.csv".decode('gbk') )


def get_data():
    list_username = []
    filename_input = filename_original
    department_to_filter = var_char_combox_department.get()
    filename_output = os.path.join(dir_filename_display, "������֤������%s����ר��������ͳ��.csv".decode('gbk') % department_to_filter)
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
    dataframe_data = pandas.DataFrame(data_display).T
    dataframe_data.to_csv(filename_output, encoding='gbk')
    tkMessageBox.showinfo('��ʾ'.decode('gbk'), '����%s�Ľ���Ѿ����ɣ���ȥ%s·���鿴.����������������������˳���ť�˳�����'.decode('gbk') % (department_to_filter, dir_filename_display))

frame_top = Tkinter.Frame(root, height=20)
frame_top.pack(side=Tkinter.TOP)

Tkinter.Label(frame_top, text='��������ѡ����Ҫ����Ĵ���'.decode('gbk'), bg='Red').pack(side=Tkinter.TOP)
box_set_department = ttk.Combobox(frame_top, textvariable=var_char_combox_department,
                                  values=['����һ��'.decode('gbk'), '���Զ���'.decode('gbk'), '��������'.decode('gbk'),
                                          '�����Ĵ�'.decode('gbk'), '�����崦'.decode('gbk'), '��������'.decode('gbk')], width=30)

box_set_department.pack(side=Tkinter.BOTTOM)

frame_middle = Tkinter.Frame(root, height=20)
frame_middle.pack()
frame_middle_top = Tkinter.Frame(frame_middle, height=40)
frame_middle_top.pack()
frame_middle_bottom = Tkinter.Frame(frame_middle, height=20)
frame_middle_bottom.pack()
Tkinter.Label(frame_middle_top, text='��������ѡ����Ҫ�����ר���ļ�'.decode('gbk'), bg='Red').pack()

Tkinter.Entry(frame_middle_bottom, textvariable=var_char_entry_filename_need_filter, width=40).pack(side=Tkinter.LEFT)
Tkinter.Button(frame_middle_bottom, text='ѡ���ļ�'.decode('gbk'), command=get_filename, width=20).pack(side=Tkinter.RIGHT)

frame_middle_2 = Tkinter.Frame(root, height=20)
frame_middle_2.pack()
frame_middle_top_2 = Tkinter.Frame(frame_middle_2, height=40)
frame_middle_top_2.pack()
frame_middle_bottom_2 = Tkinter.Frame(frame_middle_2, height=20)
frame_middle_bottom_2.pack()
Tkinter.Label(frame_middle_top_2, text='��������ѡ��������ŵ�λ��'.decode('gbk'), bg='Red').pack()
Tkinter.Entry(frame_middle_bottom_2, textvariable=var_char_entry_filename_after_filter, width=40).pack(side=Tkinter.LEFT)
Tkinter.Button(frame_middle_bottom_2, text='ѡ���ļ�'.decode('gbk'), command=set_filename, width=20).pack(side=Tkinter.RIGHT)


frame_bottom = Tkinter.Frame(root, height=20)
frame_bottom.pack()

Tkinter.Button(frame_bottom, text='GO'.decode('gbk'), width=20, command=get_data).pack(side=Tkinter.LEFT)
Tkinter.Button(frame_bottom, text='�˳�'.decode('gbk'), width=20, command=root.destroy).pack(side=Tkinter.LEFT)

Tkinter.mainloop()
