#!/bin/usr/env python
# -*- coding:cp936 -*-

"""这个工具的作用是过滤部门的专利总览表，然后统计出g各处个人的专利完成情况。分为四列：发明提交、发明受理、实用新型提交、实用新型受理。"""

import xlsxwriter
import time
import os
import wx
import wx.xrc
import xlrd
import re

DisplayFilename = wx.TextCtrl
DisplayResultDir = wx.TextCtrl
filename_original_zonglan = unicode()
filename_original_shouli = unicode()
filename_allname = unicode()
dir_filename_display = unicode()

TeamLeader = ['贾岛'.decode('gbk'), '潘霖'.decode('gbk'), '韩琳琳'.decode('gbk'), '苗永威'.decode('gbk'),
              '史沛玉'.decode('gbk'), '杨文清'.decode('gbk'), '伯绍文'.decode('gbk'), '迟江波'.decode('gbk'), '李永亮'.decode('gbk'),
              '曹翔'.decode('gbk')]

Member0 = ['贾岛'.decode('gbk'), '李光达'.decode('gbk'), '刘茂峰'.decode('gbk'), '范鹏飞'.decode('gbk'), '谭静静'.decode('gbk'),
           '张文珂'.decode('gbk'), '代如静'.decode('gbk')]
Member1 = ['潘霖'.decode('gbk'), '刘博'.decode('gbk'), '黄翼'.decode('gbk'), '董喜燕'.decode('gbk'), '郝良晟'.decode('gbk')]
Member2 = ['韩琳琳'.decode('gbk'), '林海'.decode('gbk'), '曹加峰'.decode('gbk'), '张行武'.decode('gbk'), '李建波'.decode('gbk'),
           '冯晓洁'.decode('gbk')]
Member3 = ['苗永威'.decode('gbk'), '闫硕'.decode('gbk'), '刘瑞雪'.decode('gbk'), '王云鹏'.decode('gbk'), '卢正超'.decode('gbk')]
Member4 = ['史沛玉'.decode('gbk'), '张超'.decode('gbk'), '张锟'.decode('gbk'), '刘智刚'.decode('gbk'), '巩祥文'.decode('gbk'),
           '孙玉超'.decode('gbk'), '韩超'.decode('gbk'), '徐伟超'.decode('gbk'), '赵盛'.decode('gbk'), '王建刚'.decode('gbk'),
           '高莹'.decode('gbk'), '王旭林'.decode('gbk'), '杨惠'.decode('gbk'), '程佳佳'.decode('gbk'), ]
Member5 = ['杨文清'.decode('gbk'), '贠雄斌'.decode('gbk'), '孙薇'.decode('gbk'), '李静'.decode('gbk'), '杨永峰'.decode('gbk')]
Member6 = ['伯绍文'.decode('gbk'), '李波'.decode('gbk'), '刘东伟'.decode('gbk'), '吴培琴'.decode('gbk'), '武秋星'.decode('gbk'),
           '胥志泉'.decode('gbk'), '赵召'.decode('gbk'), '李壮'.decode('gbk'), '李俊卿'.decode('gbk'), '张日洪'.decode('gbk')]
Member7 = ['迟江波'.decode('gbk'), '刘浩君'.decode('gbk'), '李彦华'.decode('gbk'), '韩燕燕'.decode('gbk'),
           '梁恒勋'.decode('gbk'), '黄锦盛'.decode('gbk'), '王晓明'.decode('gbk'), '刘学艳'.decode('gbk')]
Member8 = ['李永亮'.decode('gbk'), '李丹'.decode('gbk'), '兰太顺'.decode('gbk')]
Member9 = ['曹翔'.decode('gbk'), '周志超'.decode('gbk')]

TitleItem = ['组长名'.decode('gbk'), '组员名'.decode('gbk'), '发明提交数量'.decode('gbk'), '发明受理数量'.decode('gbk'),
             '实用新型提交数量'.decode('gbk'), '实用新型受理数量'.decode('gbk')]


class InventionFilterTeamThree(wx.Frame):
    def __init__(self, parent):
        wx.Frame.__init__(self, parent, id=wx.ID_ANY, title=u"测试三处专利结果过滤工具", pos=wx.DefaultPosition,
                          size=wx.Size(387, 355), style=wx.CAPTION | wx.RESIZE_BORDER | wx.TAB_TRAVERSAL)

        self.SetSizeHints(wx.DefaultSize, wx.DefaultSize)
        self.SetForegroundColour(wx.Colour(255, 255, 0))
        self.SetBackgroundColour(wx.Colour(72, 220, 35))

        bSizer1 = wx.BoxSizer(wx.VERTICAL)

        bSizer3 = wx.BoxSizer(wx.VERTICAL)

        self.text_1 = wx.StaticText(self, wx.ID_ANY, u"请在如下选择需要处理的专利总览文件", wx.DefaultPosition, wx.DefaultSize, 0)
        self.text_1.Wrap(-1)
        self.text_1.SetForegroundColour(wx.Colour(0, 0, 0))
        self.text_1.SetBackgroundColour(wx.Colour(255, 0, 0))

        bSizer3.Add(self.text_1, 0, wx.ALIGN_CENTER_HORIZONTAL | wx.ALL, 5)

        bSizer1.Add(bSizer3, 0, wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer2 = wx.BoxSizer(wx.HORIZONTAL)

        self.textctl_zonglan_filename = wx.TextCtrl(self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition,
                                                    wx.DefaultSize,
                                                    0)
        bSizer2.Add(self.textctl_zonglan_filename, 1, wx.ALL, 5)

        self.ButtonChoseFile = wx.Button(self, wx.ID_ANY, u"选择专利总览文件", wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer2.Add(self.ButtonChoseFile, 0, wx.ALIGN_RIGHT | wx.ALL, 5)

        bSizer1.Add(bSizer2, 0, wx.EXPAND, 5)

        bSizer9 = wx.BoxSizer(wx.VERTICAL)

        self.m_staticText4 = wx.StaticText(self, wx.ID_ANY, u"请在如下选择需要处理的专利受理文件", wx.DefaultPosition,
                                           wx.DefaultSize, 0)
        self.m_staticText4.Wrap(-1)
        self.m_staticText4.SetForegroundColour(wx.Colour(0, 0, 0))
        self.m_staticText4.SetBackgroundColour(wx.Colour(255, 0, 0))

        bSizer9.Add(self.m_staticText4, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer1.Add(bSizer9, 0, wx.EXPAND, 5)

        bSizer10 = wx.BoxSizer(wx.HORIZONTAL)

        self.textctl_shouli_filename = wx.TextCtrl(self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition,
                                                   wx.DefaultSize,
                                                   0)
        bSizer10.Add(self.textctl_shouli_filename, 1, wx.ALL, 5)

        self.m_button6 = wx.Button(self, wx.ID_ANY, u"选择专利受理文件", wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer10.Add(self.m_button6, 0, wx.ALL, 5)

        bSizer1.Add(bSizer10, 0, wx.EXPAND, 5)

        bSizer71 = wx.BoxSizer(wx.VERTICAL)

        self.m_staticText3 = wx.StaticText(self, wx.ID_ANY, u"请在如下选择人处别", wx.DefaultPosition, wx.DefaultSize, 0)
        self.m_staticText3.Wrap(-1)
        self.m_staticText3.SetForegroundColour(wx.Colour(0, 0, 0))
        self.m_staticText3.SetBackgroundColour(wx.Colour(255, 0, 0))

        bSizer71.Add(self.m_staticText3, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer1.Add(bSizer71, 0, wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer81 = wx.BoxSizer(wx.HORIZONTAL)

        combobox_teamChoices = [u"测试三处"]
        self.combobox_team = wx.ComboBox(self, wx.ID_ANY, u"测试三处", wx.DefaultPosition, wx.DefaultSize,
                                         combobox_teamChoices, 0)
        self.combobox_team.SetSelection(0)
        bSizer81.Add(self.combobox_team, 1, wx.ALL, 5)

        bSizer1.Add(bSizer81, 0, wx.EXPAND, 5)

        bSizer4 = wx.BoxSizer(wx.VERTICAL)

        self.text_2 = wx.StaticText(self, wx.ID_ANY, u"请在如下选择输出处理结果的目录", wx.DefaultPosition, wx.DefaultSize, 0)
        self.text_2.Wrap(-1)
        self.text_2.SetForegroundColour(wx.Colour(0, 0, 0))
        self.text_2.SetBackgroundColour(wx.Colour(255, 0, 0))

        bSizer4.Add(self.text_2, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer1.Add(bSizer4, 0, wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer7 = wx.BoxSizer(wx.HORIZONTAL)

        self.text_dislpay_result = wx.TextCtrl(self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize,
                                               0)
        bSizer7.Add(self.text_dislpay_result, 1, wx.ALL, 5)

        self.ButtonChoseDir = wx.Button(self, wx.ID_ANY, u"选择目录", wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer7.Add(self.ButtonChoseDir, 0, wx.ALIGN_RIGHT | wx.ALL, 5)

        bSizer1.Add(bSizer7, 0, wx.EXPAND, 5)

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
        self.ButtonChoseFile.Bind(wx.EVT_BUTTON, self.get_zonglan_filename)
        self.m_button6.Bind(wx.EVT_BUTTON, self.get_shouli_filename)
        self.ButtonChoseDir.Bind(wx.EVT_BUTTON, self.set_filename)
        self.button_go.Bind(wx.EVT_BUTTON, self.get_data)
        self.button_exit.Bind(wx.EVT_BUTTON, self.close)

    def __del__(self):
        pass

    def close(self, event):
        self.Close()

    # Virtual event handlers, overide them in your derived class
    def get_zonglan_filename(self, event):
        global filename_original_zonglan
        filename_invention_dialog = wx.FileDialog(self, message="选择专利总览文件".decode('gbk'), defaultDir=os.getcwd(),
                                                  defaultFile="")
        if filename_invention_dialog.ShowModal() == wx.ID_OK:
            filename_invention = filename_invention_dialog.GetPath()
            self.textctl_zonglan_filename.SetValue(filename_invention)
            filename_original_zonglan = filename_invention
            filename_invention_dialog.Destroy()

    def get_shouli_filename(self, event):
        global filename_original_shouli
        filename_invention_dialog = wx.FileDialog(self, message="选择专利受理文件".decode('gbk'), defaultDir=os.getcwd(),
                                                  defaultFile="")
        if filename_invention_dialog.ShowModal() == wx.ID_OK:
            filename_invention = filename_invention_dialog.GetPath()
            self.textctl_shouli_filename.SetValue(filename_invention)
            filename_original_shouli = filename_invention
            filename_invention_dialog.Destroy()

    def get_allname(self, event):
        global filename_allname
        filename_invention_dialog = wx.FileDialog(self, message="选择人员名单文件".decode('gbk'), defaultDir=os.getcwd(),
                                                  defaultFile="")
        if filename_invention_dialog.ShowModal() == wx.ID_OK:
            all_name = filename_invention_dialog.GetPath()
            self.textctrl_allname.SetValue(all_name)
            filename_allname = all_name
            filename_invention_dialog.Destroy()

    def set_filename(self, event):
        global dir_filename_display
        dir_filename_display_dialog = wx.DirDialog(self, message="选择存储目录".decode('gbk'), style=wx.DD_DEFAULT_STYLE)
        if dir_filename_display_dialog.ShowModal() == wx.ID_OK:
            dir_filename_display = dir_filename_display_dialog.GetPath()
            #            .replace('/', '\\')
            self.text_dislpay_result.SetValue(dir_filename_display)
            dir_filename_display_dialog.Destroy()

    def get_data(self, event):
        print time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
        length_team = len(TeamLeader)
        department_to_filter = self.combobox_team.GetValue()
        timestamp = time.strftime('%Y%m%d', time.localtime())
        filename_output = os.path.join(dir_filename_display,
                                       "%s各人专利完成情况统计-%s.xlsx".decode('gbk') % (department_to_filter, timestamp))
        WorkBook = xlsxwriter.Workbook(filename_output)
        SheetOne = WorkBook.add_worksheet('%s各人专利完成情况统计'.decode('gbk') % department_to_filter)
        format_out = WorkBook.add_format()
        format_out.set_border(1)
        sum_line = 0
        ListUsername = []
        for i in range(0, length_team):
            sum_line += len(globals()['Member' + str(i)])
        for i in range(0, len(TitleItem)):
            SheetOne.write(0, i, TitleItem[i], format_out)
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
                    SheetOne.write(i, 1, (globals()['Member' + str(j)])[k], format_out)
                    ListUsername.append((globals()['Member' + str(j)])[k])
                    i += 1
        data_display = {}
        list_status = ["撰写通过".decode('gbk')]
        list_except = ["待决定".decode('gbk'), "撰写驳回".decode('gbk')]
        for name in ListUsername:
            data_display['%s' % name] = {}
            data_display['%s' % name]['发明提交数量'.decode('gbk')] = 0
            data_display['%s' % name]['发明受理数量'.decode('gbk')] = 0
            data_display['%s' % name]['实用新型提交数量'.decode('gbk')] = 0
            data_display['%s' % name]['实用新型受理数量'.decode('gbk')] = 0
        file_name = xlrd.open_workbook(filename_original_zonglan, encoding_override='cp936')
        sheet_filter_one = file_name.sheet_by_index(0)
        total_rows_one = sheet_filter_one.nrows

        sheet_filter_two = file_name.sheet_by_index(1)
        total_rows_two = sheet_filter_two.nrows

        for item_1 in range(1, total_rows_one):
            username_temp = sheet_filter_one.cell(item_1, 6).value.strip().split(";")[0]
            username = re.search(r"\D*", username_temp).group()
            type_invention = sheet_filter_one.cell(item_1, 17).value.strip().split(";")[0]
            # shouli_or_not = sheet_filter_one.cell(item_1, 8).value
            status = sheet_filter_one.cell(item_1, 23).value.strip()
            if username in ListUsername:
                if status in list_status :
                    if type_invention == '发明'.decode('gbk'):
                        data_display['%s' % username]['发明受理数量'.decode('gbk')] += 1
                    if type_invention == '新型'.decode('gbk'):
                        data_display['%s' % username]['实用新型受理数量'.decode('gbk')] += 1
                if type_invention == '发明'.decode('gbk') and status not in list_except:
                    data_display['%s' % username]['发明提交数量'.decode('gbk')] += 1
                if type_invention == '新型'.decode('gbk') and status not in list_except:
                    data_display['%s' % username]['实用新型提交数量'.decode('gbk')] += 1

        for item_2 in range(1, total_rows_two):
            username = sheet_filter_two.cell(item_2, 1).value
            type_invention = sheet_filter_two.cell(item_2, 6).value.replace(u' ', u'')
            if username in ListUsername:
                if type_invention == '发明'.decode('gbk'):
                    data_display['%s' % username]['发明受理数量'.decode('gbk')] += 1
                if type_invention == '实用新型'.decode('gbk'):
                    data_display['%s' % username]['实用新型受理数量'.decode('gbk')] += 1
        SheetOne.set_column("C:F", 15)
        i = 1

        # 处理受理统计数据，查询并统计受理的专利
        file_name_shouli = xlrd.open_workbook(filename_original_shouli, encoding_override='cp936')
        sheet_filter_shouli = file_name_shouli.sheet_by_index(0)
        total_rows_shouli = sheet_filter_shouli.nrows
        for item_shouli in range(1, total_rows_shouli):
            username_shouli = sheet_filter_shouli.cell(item_shouli, 4).value.strip()
            type_invention = sheet_filter_shouli.cell(item_shouli, 1).value.split(",")[0].strip()
            if username_shouli in ListUsername:
                if type_invention == '发明'.decode('gbk'):
                    data_display['%s' % username_shouli]['发明受理数量'.decode('gbk')] += 1
                if type_invention == '新型'.decode('gbk'):
                    data_display['%s' % username_shouli]['实用新型受理数量'.decode('gbk')] += 1

        for username in ListUsername:
            SheetOne.write(i, 2, data_display['%s' % username]['发明提交数量'.decode('gbk')])
            SheetOne.write(i, 3, data_display['%s' % username]['发明受理数量'.decode('gbk')])
            SheetOne.write(i, 4, data_display['%s' % username]['实用新型提交数量'.decode('gbk')])
            SheetOne.write(i, 5, data_display['%s' % username]['实用新型受理数量'.decode('gbk')])
            i += 1
        WorkBook.close()
        print time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
        diag_finish = wx.MessageDialog(None, '处理结果已经保存到文件%s《%s》中.如果无需其他动作，请点击退出按钮退出程序'.decode('gbk') % (
            dir_filename_display, filename_output), '提示'.decode('gbk'), wx.OK | wx.ICON_INFORMATION | wx.STAY_ON_TOP)
        diag_finish.ShowModal()


if __name__ == '__main__':
    app = wx.App()
    frame = InventionFilterTeamThree(None)
    frame.Show()
    app.MainLoop()
