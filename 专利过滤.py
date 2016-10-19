#!/bin/usr/env python
# -*- coding:cp936 -*-
"""这个工具的作用是过滤部门的专利总览表，然后统计出一处和二处个人的专利完成情况。分为四列：发明提交、发明受理、实用新型提交、实用新型受理。输入文档
必须是csv格式的，切只有一个sheet。输入文件名为了方便固定位《测试验证部专利总览.csv》。输出文件名称为《测试验证部测试一处和二处个人专利完成情况统计.csv》
"""
import pandas

list_username = []
filename_input = '测试验证部专利总览.csv'.decode('gbk')
filename_output = '测试验证部测试一处和二处个人专利完成情况统计.csv'.decode('gbk')
file_name = pandas.read_csv(filename_input, sep=',', header=1, encoding='gbk', na_filter=True, chunksize=1)
for item in file_name:
    department = item.iloc[0, 0]
    if department == '测试一处'.decode('gbk'):
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
    department = item_1.iloc[0, 0]
    username = item_1.iloc[0, 1]
    type_invention = item_1.iloc[0, 6]
    date_shouli = item_1.iloc[0, 5]

    if department == '测试一处'.decode('gbk'):
        if username in list_username:
            if type_invention == '发明'.decode('gbk'):
                data_display['%s' % username]['发明提交数量'.decode('gbk')] += 1
            if type_invention == '实用新型'.decode('gbk'):
                data_display['%s' % username]['实用新型提交数量'.decode('gbk')] += 1
            if pandas.notnull(date_shouli):
                if type_invention == '发明'.decode('gbk'):
                    data_display['%s' % username]['发明受理数量'.decode('gbk')] += 1
                if type_invention == '实用新型'.decode('gbk'):
                    data_display['%s' % username]['实用新型受理数量'.decode('gbk')] += 1
dataframe_data = pandas.DataFrame(data_display).T
dataframe_data.to_csv(filename_output, encoding='gbk')
