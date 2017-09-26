#!/bin/usr/env python
# -*- coding:cp936 -*-

"""������ߵ������ǹ��˲��ŵ�ר��������Ȼ��ͳ�Ƴ�g�������˵�ר������������Ϊ���У������ύ����������ʵ�������ύ��ʵ����������"""

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

TeamLeader = ['�ֵ�'.decode('gbk'), '����'.decode('gbk'), '������'.decode('gbk'), '������'.decode('gbk'),
              'ʷ����'.decode('gbk'), '������'.decode('gbk'), '������'.decode('gbk'), '�ٽ���'.decode('gbk'), '������'.decode('gbk'),
              '����'.decode('gbk')]

Member0 = ['�ֵ�'.decode('gbk'), '����'.decode('gbk'), '��ï��'.decode('gbk'), '������'.decode('gbk'), '̷����'.decode('gbk'),
           '������'.decode('gbk'), '���羲'.decode('gbk')]
Member1 = ['����'.decode('gbk'), '����'.decode('gbk'), '����'.decode('gbk'), '��ϲ��'.decode('gbk'), '������'.decode('gbk')]
Member2 = ['������'.decode('gbk'), '�ֺ�'.decode('gbk'), '�ܼӷ�'.decode('gbk'), '������'.decode('gbk'), '���'.decode('gbk'),
           '������'.decode('gbk')]
Member3 = ['������'.decode('gbk'), '��˶'.decode('gbk'), '����ѩ'.decode('gbk'), '������'.decode('gbk'), '¬����'.decode('gbk')]
Member4 = ['ʷ����'.decode('gbk'), '�ų�'.decode('gbk'), '���'.decode('gbk'), '���Ǹ�'.decode('gbk'), '������'.decode('gbk'),
           '����'.decode('gbk'), '����'.decode('gbk'), '��ΰ��'.decode('gbk'), '��ʢ'.decode('gbk'), '������'.decode('gbk'),
           '��Ө'.decode('gbk'), '������'.decode('gbk'), '���'.decode('gbk'), '�̼Ѽ�'.decode('gbk'), ]
Member5 = ['������'.decode('gbk'), '�O�۱�'.decode('gbk'), '��ޱ'.decode('gbk'), '�'.decode('gbk'), '������'.decode('gbk')]
Member6 = ['������'.decode('gbk'), '�'.decode('gbk'), '����ΰ'.decode('gbk'), '������'.decode('gbk'), '������'.decode('gbk'),
           '��־Ȫ'.decode('gbk'), '����'.decode('gbk'), '��׳'.decode('gbk'), '���'.decode('gbk'), '���պ�'.decode('gbk')]
Member7 = ['�ٽ���'.decode('gbk'), '���ƾ�'.decode('gbk'), '���廪'.decode('gbk'), '������'.decode('gbk'),
           '����ѫ'.decode('gbk'), '�ƽ�ʢ'.decode('gbk'), '������'.decode('gbk'), '��ѧ��'.decode('gbk')]
Member8 = ['������'.decode('gbk'), '�'.decode('gbk'), '��̫˳'.decode('gbk')]
Member9 = ['����'.decode('gbk'), '��־��'.decode('gbk')]

TitleItem = ['�鳤��'.decode('gbk'), '��Ա��'.decode('gbk'), '�����ύ����'.decode('gbk'), '������������'.decode('gbk'),
             'ʵ�������ύ����'.decode('gbk'), 'ʵ��������������'.decode('gbk')]


class InventionFilterTeamThree(wx.Frame):
    def __init__(self, parent):
        wx.Frame.__init__(self, parent, id=wx.ID_ANY, title=u"��������ר��������˹���", pos=wx.DefaultPosition,
                          size=wx.Size(387, 355), style=wx.CAPTION | wx.RESIZE_BORDER | wx.TAB_TRAVERSAL)

        self.SetSizeHints(wx.DefaultSize, wx.DefaultSize)
        self.SetForegroundColour(wx.Colour(255, 255, 0))
        self.SetBackgroundColour(wx.Colour(72, 220, 35))

        bSizer1 = wx.BoxSizer(wx.VERTICAL)

        bSizer3 = wx.BoxSizer(wx.VERTICAL)

        self.text_1 = wx.StaticText(self, wx.ID_ANY, u"��������ѡ����Ҫ�����ר�������ļ�", wx.DefaultPosition, wx.DefaultSize, 0)
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

        self.ButtonChoseFile = wx.Button(self, wx.ID_ANY, u"ѡ��ר�������ļ�", wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer2.Add(self.ButtonChoseFile, 0, wx.ALIGN_RIGHT | wx.ALL, 5)

        bSizer1.Add(bSizer2, 0, wx.EXPAND, 5)

        bSizer9 = wx.BoxSizer(wx.VERTICAL)

        self.m_staticText4 = wx.StaticText(self, wx.ID_ANY, u"��������ѡ����Ҫ�����ר�������ļ�", wx.DefaultPosition,
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

        self.m_button6 = wx.Button(self, wx.ID_ANY, u"ѡ��ר�������ļ�", wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer10.Add(self.m_button6, 0, wx.ALL, 5)

        bSizer1.Add(bSizer10, 0, wx.EXPAND, 5)

        bSizer71 = wx.BoxSizer(wx.VERTICAL)

        self.m_staticText3 = wx.StaticText(self, wx.ID_ANY, u"��������ѡ���˴���", wx.DefaultPosition, wx.DefaultSize, 0)
        self.m_staticText3.Wrap(-1)
        self.m_staticText3.SetForegroundColour(wx.Colour(0, 0, 0))
        self.m_staticText3.SetBackgroundColour(wx.Colour(255, 0, 0))

        bSizer71.Add(self.m_staticText3, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer1.Add(bSizer71, 0, wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer81 = wx.BoxSizer(wx.HORIZONTAL)

        combobox_teamChoices = [u"��������"]
        self.combobox_team = wx.ComboBox(self, wx.ID_ANY, u"��������", wx.DefaultPosition, wx.DefaultSize,
                                         combobox_teamChoices, 0)
        self.combobox_team.SetSelection(0)
        bSizer81.Add(self.combobox_team, 1, wx.ALL, 5)

        bSizer1.Add(bSizer81, 0, wx.EXPAND, 5)

        bSizer4 = wx.BoxSizer(wx.VERTICAL)

        self.text_2 = wx.StaticText(self, wx.ID_ANY, u"��������ѡ�������������Ŀ¼", wx.DefaultPosition, wx.DefaultSize, 0)
        self.text_2.Wrap(-1)
        self.text_2.SetForegroundColour(wx.Colour(0, 0, 0))
        self.text_2.SetBackgroundColour(wx.Colour(255, 0, 0))

        bSizer4.Add(self.text_2, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer1.Add(bSizer4, 0, wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer7 = wx.BoxSizer(wx.HORIZONTAL)

        self.text_dislpay_result = wx.TextCtrl(self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize,
                                               0)
        bSizer7.Add(self.text_dislpay_result, 1, wx.ALL, 5)

        self.ButtonChoseDir = wx.Button(self, wx.ID_ANY, u"ѡ��Ŀ¼", wx.DefaultPosition, wx.DefaultSize, 0)
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
        filename_invention_dialog = wx.FileDialog(self, message="ѡ��ר�������ļ�".decode('gbk'), defaultDir=os.getcwd(),
                                                  defaultFile="")
        if filename_invention_dialog.ShowModal() == wx.ID_OK:
            filename_invention = filename_invention_dialog.GetPath()
            self.textctl_zonglan_filename.SetValue(filename_invention)
            filename_original_zonglan = filename_invention
            filename_invention_dialog.Destroy()

    def get_shouli_filename(self, event):
        global filename_original_shouli
        filename_invention_dialog = wx.FileDialog(self, message="ѡ��ר�������ļ�".decode('gbk'), defaultDir=os.getcwd(),
                                                  defaultFile="")
        if filename_invention_dialog.ShowModal() == wx.ID_OK:
            filename_invention = filename_invention_dialog.GetPath()
            self.textctl_shouli_filename.SetValue(filename_invention)
            filename_original_shouli = filename_invention
            filename_invention_dialog.Destroy()

    def get_allname(self, event):
        global filename_allname
        filename_invention_dialog = wx.FileDialog(self, message="ѡ����Ա�����ļ�".decode('gbk'), defaultDir=os.getcwd(),
                                                  defaultFile="")
        if filename_invention_dialog.ShowModal() == wx.ID_OK:
            all_name = filename_invention_dialog.GetPath()
            self.textctrl_allname.SetValue(all_name)
            filename_allname = all_name
            filename_invention_dialog.Destroy()

    def set_filename(self, event):
        global dir_filename_display
        dir_filename_display_dialog = wx.DirDialog(self, message="ѡ��洢Ŀ¼".decode('gbk'), style=wx.DD_DEFAULT_STYLE)
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
                                       "%s����ר��������ͳ��-%s.xlsx".decode('gbk') % (department_to_filter, timestamp))
        WorkBook = xlsxwriter.Workbook(filename_output)
        SheetOne = WorkBook.add_worksheet('%s����ר��������ͳ��'.decode('gbk') % department_to_filter)
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
        list_status = ["׫дͨ��".decode('gbk')]
        list_except = ["������".decode('gbk'), "׫д����".decode('gbk')]
        for name in ListUsername:
            data_display['%s' % name] = {}
            data_display['%s' % name]['�����ύ����'.decode('gbk')] = 0
            data_display['%s' % name]['������������'.decode('gbk')] = 0
            data_display['%s' % name]['ʵ�������ύ����'.decode('gbk')] = 0
            data_display['%s' % name]['ʵ��������������'.decode('gbk')] = 0
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
                    if type_invention == '����'.decode('gbk'):
                        data_display['%s' % username]['������������'.decode('gbk')] += 1
                    if type_invention == '����'.decode('gbk'):
                        data_display['%s' % username]['ʵ��������������'.decode('gbk')] += 1
                if type_invention == '����'.decode('gbk') and status not in list_except:
                    data_display['%s' % username]['�����ύ����'.decode('gbk')] += 1
                if type_invention == '����'.decode('gbk') and status not in list_except:
                    data_display['%s' % username]['ʵ�������ύ����'.decode('gbk')] += 1

        for item_2 in range(1, total_rows_two):
            username = sheet_filter_two.cell(item_2, 1).value
            type_invention = sheet_filter_two.cell(item_2, 6).value.replace(u' ', u'')
            if username in ListUsername:
                if type_invention == '����'.decode('gbk'):
                    data_display['%s' % username]['������������'.decode('gbk')] += 1
                if type_invention == 'ʵ������'.decode('gbk'):
                    data_display['%s' % username]['ʵ��������������'.decode('gbk')] += 1
        SheetOne.set_column("C:F", 15)
        i = 1

        # ��������ͳ�����ݣ���ѯ��ͳ�������ר��
        file_name_shouli = xlrd.open_workbook(filename_original_shouli, encoding_override='cp936')
        sheet_filter_shouli = file_name_shouli.sheet_by_index(0)
        total_rows_shouli = sheet_filter_shouli.nrows
        for item_shouli in range(1, total_rows_shouli):
            username_shouli = sheet_filter_shouli.cell(item_shouli, 4).value.strip()
            type_invention = sheet_filter_shouli.cell(item_shouli, 1).value.split(",")[0].strip()
            if username_shouli in ListUsername:
                if type_invention == '����'.decode('gbk'):
                    data_display['%s' % username_shouli]['������������'.decode('gbk')] += 1
                if type_invention == '����'.decode('gbk'):
                    data_display['%s' % username_shouli]['ʵ��������������'.decode('gbk')] += 1

        for username in ListUsername:
            SheetOne.write(i, 2, data_display['%s' % username]['�����ύ����'.decode('gbk')])
            SheetOne.write(i, 3, data_display['%s' % username]['������������'.decode('gbk')])
            SheetOne.write(i, 4, data_display['%s' % username]['ʵ�������ύ����'.decode('gbk')])
            SheetOne.write(i, 5, data_display['%s' % username]['ʵ��������������'.decode('gbk')])
            i += 1
        WorkBook.close()
        print time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
        diag_finish = wx.MessageDialog(None, '�������Ѿ����浽�ļ�%s��%s����.����������������������˳���ť�˳�����'.decode('gbk') % (
            dir_filename_display, filename_output), '��ʾ'.decode('gbk'), wx.OK | wx.ICON_INFORMATION | wx.STAY_ON_TOP)
        diag_finish.ShowModal()


if __name__ == '__main__':
    app = wx.App()
    frame = InventionFilterTeamThree(None)
    frame.Show()
    app.MainLoop()
