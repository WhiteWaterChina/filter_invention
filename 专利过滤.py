#!/bin/usr/env python
# -*- coding:cp936 -*-
"""������ߵ������ǹ��˲��ŵ�ר��������Ȼ��ͳ�Ƴ�һ���Ͷ������˵�ר������������Ϊ���У������ύ����������ʵ�������ύ��ʵ���������������ĵ�
������csv��ʽ�ģ���ֻ��һ��sheet�������ļ���Ϊ�˷���̶�λ��������֤��ר������.csv��������ļ�����Ϊ��������֤������һ���Ͷ�������ר��������ͳ��.csv��
"""
import pandas

list_username = []
filename_input = '������֤��ר������.csv'.decode('gbk')
filename_output = '������֤������һ���Ͷ�������ר��������ͳ��.csv'.decode('gbk')
file_name = pandas.read_csv(filename_input, sep=',', header=1, encoding='gbk', na_filter=True, chunksize=1)
for item in file_name:
    department = item.iloc[0, 0]
    if department == '����һ��'.decode('gbk'):
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
    department = item_1.iloc[0, 0]
    username = item_1.iloc[0, 1]
    type_invention = item_1.iloc[0, 6]
    date_shouli = item_1.iloc[0, 5]

    if department == '����һ��'.decode('gbk'):
        if username in list_username:
            if type_invention == '����'.decode('gbk'):
                data_display['%s' % username]['�����ύ����'.decode('gbk')] += 1
            if type_invention == 'ʵ������'.decode('gbk'):
                data_display['%s' % username]['ʵ�������ύ����'.decode('gbk')] += 1
            if pandas.notnull(date_shouli):
                if type_invention == '����'.decode('gbk'):
                    data_display['%s' % username]['������������'.decode('gbk')] += 1
                if type_invention == 'ʵ������'.decode('gbk'):
                    data_display['%s' % username]['ʵ��������������'.decode('gbk')] += 1
dataframe_data = pandas.DataFrame(data_display).T
dataframe_data.to_csv(filename_output, encoding='gbk')
