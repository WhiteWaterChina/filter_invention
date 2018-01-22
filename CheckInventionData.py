#!/bin/usr/env python
# -*- coding:cp936 -*-


import xlsxwriter
import time
import os
import wx
import wx.xrc
import xlrd
import re

faming_name_in_jishu = []
faming_sn_in_jishu = []
faming_inventor_in_jishu = []

shiyong_name_in_jishu = []
shiyong_sn_in_jishu = []
shiyong_inventor_in_jishu = []

faming_name_in_me = []
faming_sn_in_me = []
faming_inventor_in_me = []

shiyong_name_in_me = []
shiyong_sn_in_me = []
shiyong_inventor_in_me = []

other_name_in_me = []
other_sn_in_me = []
other_inventor_in_me = []

#�򿪼��������ĵ���ȡ����
file_name_jishu = xlrd.open_workbook("������������.xlsx".decode('gbk'), encoding_override='cp936')
sheet_filter_jishu = file_name_jishu.sheet_by_index(0)
total_rows_jishu = sheet_filter_jishu.nrows

#��������
for item_1 in range(1, total_rows_jishu):
    type_jishu = sheet_filter_jishu.cell(item_1, 5).value.strip()
    filename_jishu_faming = sheet_filter_jishu.cell(item_1, 4).value.strip()
    shoulinumber_jishu_faming = sheet_filter_jishu.cell(item_1, 3).value
    inventor_jishu = sheet_filter_jishu.cell(item_1,6).value

    #faming_name_in_jishu.append(filename_jishu_faming)
    #faming_sn_in_jishu.append(shoulinumber_jishu_faming)
    if type_jishu == "����".decode('gbk'):
        faming_name_in_jishu.append(filename_jishu_faming)
        faming_sn_in_jishu.append(shoulinumber_jishu_faming)
        faming_inventor_in_jishu.append(inventor_jishu)
    elif type_jishu == "ʵ������".decode('gbk'):
        shiyong_name_in_jishu.append(filename_jishu_faming)
        shiyong_sn_in_jishu.append(shoulinumber_jishu_faming)
        shiyong_inventor_in_jishu.append(inventor_jishu)
# #ʵ������
# for item_1 in range(1, total_rows_jishu_shiyong):
#     filename_jishu_shiyong = sheet_filter_jishu_shiyong.cell(item_1, 5).value.strip()
#     shoulinumber_jishu_shiyong = sheet_filter_jishu_shiyong.cell(item_1, 4).value
#     shiyong_name_in_jishu.append(filename_jishu_shiyong)
#     shiyong_sn_in_jishu.append(shoulinumber_jishu_shiyong)


#��ȡ�ҵ�����
file_name_me = xlrd.open_workbook("�ҵ�����.xlsx".decode('gbk'), encoding_override='cp936')
sheet_filter_me = file_name_me.sheet_by_index(0)
total_rows_jme = sheet_filter_me.nrows
for item_1 in range(2, total_rows_jme):
    type_invention = sheet_filter_me.cell(item_1,1).value.strip().split(",")[0]
#    print type_invention
    filename_me = sheet_filter_me.cell(item_1, 4).value.strip()
    shoulinumber_me = sheet_filter_me.cell(item_1, 6).value.strip().lower()
    if type_invention == "����".decode('gbk'):
        faming_name_in_me.append(filename_me)
        faming_sn_in_me.append(shoulinumber_me)
    elif type_invention == "����".decode('gbk'):
        shiyong_name_in_me.append(filename_me)
        shiyong_sn_in_me.append(shoulinumber_me)
    else:
        other_name_in_me.append(filename_me)
        other_sn_in_me.append(shoulinumber_me)
list_name_in_me_faming = []
list_sn_in_me_faming = []
list_name_in_me_shiyong  = []
list_sn_in_me_shiyong = []
#�Ա�����
#���ҵ������У����ǲ��ڼ����������е�
for index_me_faming, item_me_faming in enumerate(faming_name_in_me):
    if item_me_faming not in faming_name_in_jishu:
        list_name_in_me_faming.append(item_me_faming)
        list_sn_in_me_faming.append(faming_sn_in_me[index_me_faming])
for index_me_shiyong, item_me_shiyong in enumerate(shiyong_name_in_me):
    if item_me_shiyong not in shiyong_name_in_jishu:
        list_name_in_me_shiyong.append(item_me_shiyong)
        list_sn_in_me_shiyong.append(shiyong_sn_in_me[index_me_shiyong])

list_name_in_jishu_faming = []
list_sn_in_jishu_faming = []
list_name_in_jishu_shiyong = []
list_sn_in_jishu_shiyong = []
#�ڼ����������У����ǲ����ҵ������е�
for index_jishu_faming, item_jishu_faming in enumerate(faming_name_in_jishu):
    if item_jishu_faming not in faming_name_in_me:
        list_name_in_jishu_faming.append(item_jishu_faming)
        list_sn_in_jishu_faming.append(faming_sn_in_jishu[index_jishu_faming])
for index_jishu_shiyong, item_jishu_shiyong in enumerate(shiyong_name_in_jishu):
    if item_jishu_shiyong not in shiyong_name_in_me:
        list_name_in_jishu_shiyong.append(item_jishu_shiyong)
        list_sn_in_jishu_shiyong.append(shiyong_sn_in_jishu[index_jishu_shiyong])

#��ʼд���ݵ��ļ�
workbook_to_write = xlsxwriter.Workbook("�ҵ�ͳ�ƺͼ�������ͳ�Ʋ��.xlsx".decode('gbk'))

formatone = workbook_to_write.add_format()
formatone.set_border(1)
sheet_in_me_faming = workbook_to_write.add_worksheet("���ҵĵ��ǲ��ڼ��������ķ���.xlsx".decode('gbk'))
#workbook_to_write_in_me_shiyong = xlsxwriter.Workbook("���ҵĵ��ǲ��ڼ���������ʵ��.xlsx".decode('gbk'))
sheet_in_me_shiyong = workbook_to_write.add_worksheet("���ҵĵ��ǲ��ڼ���������ʵ��.xlsx".decode('gbk'))
#workbook_to_write_in_jishu_faming = xlsxwriter.Workbook("�ڼ��������ĵ��ǲ����ҵķ���.xlsx".decode('gbk'))
sheet_in_jishu_faming = workbook_to_write.add_worksheet("�ڼ��������ĵ��ǲ����ҵķ���.xlsx".decode('gbk'))
#workbook_to_write_in_jishu_shiyong = xlsxwriter.Workbook("�ڼ��������ĵ��ǲ����ҵ�ʵ��.xlsx".decode('gbk'))
sheet_in_jishu_shiyong = workbook_to_write.add_worksheet("�ڼ��������ĵ��ǲ����ҵ�ʵ��.xlsx".decode('gbk'))
sheet_in_me_other = workbook_to_write.add_worksheet("other.xlsx".decode('gbk'))


for index, item in enumerate(list_name_in_me_faming):
    sheet_in_me_faming.write(index,0, item)
    sheet_in_me_faming.write(index,1,list_sn_in_me_faming[index])
for index, item in enumerate(list_name_in_me_shiyong):
    sheet_in_me_shiyong.write(index,0, item)
    sheet_in_me_shiyong.write(index,1, list_sn_in_me_shiyong[index])
for index, item in enumerate(list_name_in_jishu_faming):
    sheet_in_jishu_faming.write(index,0, item)
    sheet_in_jishu_faming.write(index,1, list_sn_in_jishu_faming[index])
for index, item in enumerate(list_name_in_jishu_shiyong):
    sheet_in_jishu_shiyong.write(index,0, item)
    sheet_in_jishu_shiyong.write(index,1, list_sn_in_jishu_shiyong[index])
for index, item in enumerate(other_name_in_me):
    sheet_in_me_other.write(index,0, item)
    sheet_in_me_other.write(index,1, other_sn_in_me[index])
workbook_to_write.close()

# for item in list_me_other:
#     print item



