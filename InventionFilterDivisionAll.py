#!/bin/usr/env python
# -*- coding:cp936 -*-

###########################################################################
## Python code generated with wxFormBuilder (version Jun 17 2015)
## http://www.wxformbuilder.org/
##
## PLEASE DO "NOT" EDIT THIS FILE!
###########################################################################
"""这个工具的作用是过滤部门的专利总览表，然后统计出g各处个人的专利完成情况。分为四列：发明提交、发明受理、实用新型提交、实用新型受理。输入文档
必须是csv格式的，且只有一个sheet。输出文件名为《测试验证部测试个人专利完成情况统计.csv》,依据不同的处别而不同在不同的sheet。 本工具基于Python Tkinter制作图形界面。依赖详见import部分。
打包成exe格式请使用pyinstall,命令为Python pyinstaller.py -F InvertionFilter.py
"""
import pandas
import xlsxwriter
import time
import os
import wx
import wx.xrc
import xlrd


###########################################################################
## Class InventionFilterAll
###########################################################################
DisplayFilename = wx.TextCtrl
DisplayResultDir = wx.TextCtrl
filename_original = unicode()
listDivisionName=['测试一处'.decode('gbk'), '测试二处'.decode('gbk'), '测试三处'.decode('gbk'), '测试四处'.decode('gbk'), '测试五处'.decode('gbk'), '测试六处'.decode('gbk')]
listTitle = ['发明提交数量'.decode('gbk'), '发明受理数量'.decode('gbk'), '实用新型提交数量'.decode('gbk'), '实用新型受理数量'.decode('gbk')]


class InventionFilterAll(wx.Frame):
    def __init__(self, parent):
        wx.Frame.__init__(self, parent, id=wx.ID_ANY, title=u"专利结果过滤工具", pos=wx.DefaultPosition, size=wx.Size(356, 215),
                          style=wx.CAPTION | wx.RESIZE_BORDER | wx.TAB_TRAVERSAL)

        self.SetSizeHintsSz(wx.DefaultSize, wx.DefaultSize)
        self.SetForegroundColour(wx.Colour(255, 255, 0))
        self.SetBackgroundColour(wx.Colour(72, 220, 35))

        bSizer1 = wx.BoxSizer(wx.VERTICAL)

        bSizer3 = wx.BoxSizer(wx.VERTICAL)

        self.text_1 = wx.StaticText(self, wx.ID_ANY, u"请在如下选择需要处理的专利统计文件", wx.DefaultPosition, wx.DefaultSize, 0)
        self.text_1.Wrap(-1)
        self.text_1.SetForegroundColour(wx.Colour(0, 0, 0))
        self.text_1.SetBackgroundColour(wx.Colour(255, 0, 0))

        bSizer3.Add(self.text_1, 0, wx.ALIGN_CENTER_HORIZONTAL | wx.ALL, 5)

        bSizer1.Add(bSizer3, 1, wx.EXPAND, 5)

        bSizer2 = wx.BoxSizer(wx.HORIZONTAL)

        self.DisplayFilename = wx.TextCtrl(self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer2.Add(self.DisplayFilename, 1, wx.ALL, 5)

        self.ButtonChoseFile = wx.Button(self, wx.ID_ANY, u"选择文件", wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer2.Add(self.ButtonChoseFile, 0, wx.ALIGN_RIGHT | wx.ALL, 5)

        bSizer1.Add(bSizer2, 1, wx.EXPAND, 5)

        bSizer4 = wx.BoxSizer(wx.VERTICAL)

        self.text_2 = wx.StaticText(self, wx.ID_ANY, u"请在如下选择输出处理结果的目录", wx.DefaultPosition, wx.DefaultSize, 0)
        self.text_2.Wrap(-1)
        self.text_2.SetForegroundColour(wx.Colour(0, 0, 0))
        self.text_2.SetBackgroundColour(wx.Colour(255, 0, 0))

        bSizer4.Add(self.text_2, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer1.Add(bSizer4, 1, wx.EXPAND, 5)

        bSizer7 = wx.BoxSizer(wx.HORIZONTAL)

        self.DisplayResultDir = wx.TextCtrl(self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer7.Add(self.DisplayResultDir, 1, wx.ALL, 5)

        self.ButtonChoseDir = wx.Button(self, wx.ID_ANY, u"选择目录", wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer7.Add(self.ButtonChoseDir, 0, wx.ALIGN_RIGHT | wx.ALL, 5)

        bSizer1.Add(bSizer7, 1, wx.EXPAND, 5)

        bSizer8 = wx.BoxSizer(wx.HORIZONTAL)

        self.button_go = wx.Button(self, wx.ID_ANY, u"GO", wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer8.Add(self.button_go, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL | wx.EXPAND, 5)

        self.button_exit = wx.Button(self, wx.ID_ANY, u"EXIT", wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer8.Add(self.button_exit, 0, wx.ALL | wx.EXPAND, 5)

        bSizer1.Add(bSizer8, 1, wx.ALIGN_CENTER_HORIZONTAL, 5)

        self.SetSizer(bSizer1)
        self.Layout()

        self.Centre(wx.BOTH)

        # Connect Events
        self.ButtonChoseFile.Bind(wx.EVT_BUTTON, self.get_filename)
        self.ButtonChoseDir.Bind(wx.EVT_BUTTON, self.set_filename)
        self.button_go.Bind(wx.EVT_BUTTON, self.get_data)
        self.button_exit.Bind(wx.EVT_BUTTON, self.close)

    def __del__(self):
        pass

    def close(self, event):
        self.Close()

    # Virtual event handlers, overide them in your derived class
    def get_filename(self, event):
        global filename_original
        filename_invention_dialog = wx.FileDialog(self, message=u"选择专利文件", defaultDir=os.getcwd(), defaultFile="", style=wx.OPEN | wx.MULTIPLE)
        if filename_invention_dialog.ShowModal() == wx.ID_OK:
            filename_invention = filename_invention_dialog.GetPath()
            self.DisplayFilename.SetValue(filename_invention)
            filename_original = filename_invention
            filename_invention_dialog.Destroy()

    def set_filename(self, event):
        global dir_filename_display
        dir_filename_display_dialog = wx.DirDialog(self, message=u"选择存储目录",style=wx.DD_DEFAULT_STYLE)
        if dir_filename_display_dialog.ShowModal() == wx.ID_OK:
            dir_filename_display = dir_filename_display_dialog.GetPath()
#            .replace('/', '\\')
            self.DisplayResultDir.SetValue(dir_filename_display)
            dir_filename_display_dialog.Destroy()

    def get_data(self, event):
        print time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
        filename_input = filename_original
        timestamp = time.strftime('%Y%m%d', time.localtime())
        filename_display = "测试验证部个人专利完成情况统计-%s.xlsx".decode('gbk') % timestamp
        filename_final = os.path.join(dir_filename_display, filename_display)
        WorkBook = xlsxwriter.Workbook(filename_final)
        formatOne = WorkBook.add_format()
        formatOne.set_border(1)
        for item_sheet_one in listDivisionName:
            WorkBook.add_worksheet(item_sheet_one)
        for department_to_filter in listDivisionName:
            sheet_now = WorkBook.get_worksheet_by_name(department_to_filter)
            sheet_now.set_column('B:E', 15)
            for i in range(1, len(listTitle) + 1):
                sheet_now.write(0, i, listTitle[i - 1], formatOne)
            list_username = []

            file_name = xlrd.open_workbook(filename_input, encoding_override='cp936')
            sheet_filter_one = file_name.sheet_by_index(0)
            total_rows_one = sheet_filter_one.nrows
            for item_sheet_one in range(1, total_rows_one):
                department = sheet_filter_one.cell(item_sheet_one, 3).value.replace(u' ', u'')
                if department == department_to_filter:
                    username = sheet_filter_one.cell(item_sheet_one, 6).value
                    if username not in list_username:
                        list_username.append(username)

            sheet_filter_two = file_name.sheet_by_index(1)
            total_rows_two = sheet_filter_two.nrows
            for item_sheet_two in range(1, total_rows_two):
                department = sheet_filter_two.cell(item_sheet_two, 0).value.replace(u' ', u'')
                if department == department_to_filter:
                    username = sheet_filter_two.cell(item_sheet_two, 1).value
                    if username not in list_username:
                        list_username.append(username)

            data_display = {}
            for name in list_username:
                data_display['%s' % name] = {}
                data_display['%s' % name]['发明提交数量'.decode('gbk')] = 0
                data_display['%s' % name]['发明受理数量'.decode('gbk')] = 0
                data_display['%s' % name]['实用新型提交数量'.decode('gbk')] = 0
                data_display['%s' % name]['实用新型受理数量'.decode('gbk')] = 0
            for item_1 in range(1, total_rows_one):
                department = sheet_filter_one.cell(item_1, 3).value.replace(u' ', u'')
                username = sheet_filter_one.cell(item_1, 6).value
                type_invention = sheet_filter_one.cell(item_1, 4).value.replace(u' ', u'')
                shouli_or_not =sheet_filter_one.cell(item_1, 8).value
                if department == department_to_filter:
                    if username in list_username:
                        if shouli_or_not != 'None':
                            if type_invention == '发明'.decode('gbk'):
                                data_display['%s' % username]['发明受理数量'.decode('gbk')] += 1
                            if type_invention == '新型'.decode('gbk'):
                                data_display['%s' % username]['实用新型受理数量'.decode('gbk')] += 1
                        if type_invention == '发明'.decode('gbk'):
                            data_display['%s' % username]['发明提交数量'.decode('gbk')] += 1
                        if type_invention == '新型'.decode('gbk'):
                            data_display['%s' % username]['实用新型提交数量'.decode('gbk')] += 1

            for item_2 in range(1, total_rows_two):
                department = sheet_filter_two.cell(item_2, 0).value.replace(u' ', u'')
                username = sheet_filter_two.cell(item_2, 1).value
                type_invention = sheet_filter_two.cell(item_2, 6).value.replace(u' ', u'')
                if department == department_to_filter:
                    if username in list_username:
                        if type_invention == '发明'.decode('gbk'):
                            data_display['%s' % username]['发明受理数量'.decode('gbk')] += 1
                        if type_invention == '实用新型'.decode('gbk'):
                            data_display['%s' % username]['实用新型受理数量'.decode('gbk')] += 1

            for index, item_sheet_one in enumerate(list_username):
                line_count = index + 1
                sheet_now.write(line_count, 0, item_sheet_one, formatOne)
                sheet_now.write(line_count, 1, data_display['%s' % item_sheet_one]['发明提交数量'.decode('gbk')], formatOne)
                sheet_now.write(line_count, 2, data_display['%s' % item_sheet_one]['发明受理数量'.decode('gbk')], formatOne)
                sheet_now.write(line_count, 3, data_display['%s' % item_sheet_one]['实用新型提交数量'.decode('gbk')], formatOne)
                sheet_now.write(line_count, 4, data_display['%s' % item_sheet_one]['实用新型受理数量'.decode('gbk')], formatOne)

        print time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
        diag_finish = wx.MessageDialog(None, '处理结果已经保存到文件%s《%s》中.如果无需其他动作，请点击退出按钮退出程序'.decode('gbk')  % (dir_filename_display, filename_display), '提示'.decode('gbk'), wx.OK |wx.ICON_INFORMATION | wx.STAY_ON_TOP  )
        diag_finish.ShowModal()
        WorkBook.close()


if __name__ == '__main__':
    app = wx.App()
    frame = InventionFilterAll(None)
    frame.Show()
    app.MainLoop()
