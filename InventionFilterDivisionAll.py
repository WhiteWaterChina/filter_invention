#!/bin/usr/env python
# -*- coding:cp936 -*-

"""������ߵ������ǹ��˲��ŵ�ר��������Ȼ��ͳ�Ƴ�g�������˵�ר������������Ϊ���У������ύ����������ʵ�������ύ��ʵ����������"""

import xlsxwriter
import time
import os
import wx
import wx.xrc
import xlrd

DisplayFilename = wx.TextCtrl
DisplayResultDir = wx.TextCtrl
filename_original_zonglan = unicode()
filename_original_shouli = unicode()
filename_allname = unicode()
dir_filename_display = unicode()

listDivisionName = ['����һ��'.decode('gbk'), '���Զ���'.decode('gbk'), '��������'.decode('gbk'), '�����Ĵ�'.decode('gbk'),
                    '�����崦'.decode('gbk'), '��������'.decode('gbk'), '�����ߴ�'.decode('gbk')]
listTitle = ['�����ύ����'.decode('gbk'), '������������'.decode('gbk'), 'ʵ�������ύ����'.decode('gbk'), 'ʵ��������������'.decode('gbk')]


# namelist = {}
# for index_chu, item_chu in enumerate(listDivisionName):
#     namelist["%s" % item_chu ] = {}


class InventionFilterAll(wx.Frame):
    def __init__(self, parent):
        wx.Frame.__init__(self, parent, id=wx.ID_ANY, title=u"ר��������˹���", pos=wx.DefaultPosition, size=wx.Size(387, 355),
                          style=wx.CAPTION | wx.RESIZE_BORDER | wx.TAB_TRAVERSAL)

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

        self.textctl_zonglan_filename = wx.TextCtrl(self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize,
                                                    0)
        bSizer2.Add(self.textctl_zonglan_filename, 1, wx.ALL, 5)

        self.ButtonChoseFile = wx.Button(self, wx.ID_ANY, u"ѡ��ר�������ļ�", wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer2.Add(self.ButtonChoseFile, 0, wx.ALIGN_RIGHT | wx.ALL, 5)

        bSizer1.Add(bSizer2, 0, wx.EXPAND, 5)

        bSizer9 = wx.BoxSizer(wx.VERTICAL)

        self.m_staticText4 = wx.StaticText(self, wx.ID_ANY, u"��������ѡ����Ҫ�����ר�������ļ�", wx.DefaultPosition, wx.DefaultSize, 0)
        self.m_staticText4.Wrap(-1)
        self.m_staticText4.SetForegroundColour(wx.Colour(0, 0, 0))
        self.m_staticText4.SetBackgroundColour(wx.Colour(255, 0, 0))

        bSizer9.Add(self.m_staticText4, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer1.Add(bSizer9, 0, wx.EXPAND, 5)

        bSizer10 = wx.BoxSizer(wx.HORIZONTAL)

        self.textctl_shouli_filename = wx.TextCtrl(self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize,
                                                   0)
        bSizer10.Add(self.textctl_shouli_filename, 1, wx.ALL, 5)

        self.m_button6 = wx.Button(self, wx.ID_ANY, u"ѡ��ר�������ļ�", wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer10.Add(self.m_button6, 0, wx.ALL, 5)

        bSizer1.Add(bSizer10, 0, wx.EXPAND, 5)

        bSizer71 = wx.BoxSizer(wx.VERTICAL)

        self.m_staticText3 = wx.StaticText(self, wx.ID_ANY, u"��������ѡ����Ա�����ļ�", wx.DefaultPosition, wx.DefaultSize, 0)
        self.m_staticText3.Wrap(-1)
        self.m_staticText3.SetForegroundColour(wx.Colour(0, 0, 0))
        self.m_staticText3.SetBackgroundColour(wx.Colour(255, 0, 0))

        bSizer71.Add(self.m_staticText3, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer1.Add(bSizer71, 0, wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer81 = wx.BoxSizer(wx.HORIZONTAL)

        self.textctrl_allname = wx.TextCtrl(self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer81.Add(self.textctrl_allname, 1, wx.ALL, 5)

        self.buttonchosename = wx.Button(self, wx.ID_ANY, u"ѡ���ļ�", wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer81.Add(self.buttonchosename, 0, wx.ALL, 5)

        bSizer1.Add(bSizer81, 0, wx.EXPAND, 5)

        bSizer4 = wx.BoxSizer(wx.VERTICAL)

        self.text_2 = wx.StaticText(self, wx.ID_ANY, u"��������ѡ�������������Ŀ¼", wx.DefaultPosition, wx.DefaultSize, 0)
        self.text_2.Wrap(-1)
        self.text_2.SetForegroundColour(wx.Colour(0, 0, 0))
        self.text_2.SetBackgroundColour(wx.Colour(255, 0, 0))

        bSizer4.Add(self.text_2, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer1.Add(bSizer4, 0, wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer7 = wx.BoxSizer(wx.HORIZONTAL)

        self.text_dislpay_result = wx.TextCtrl(self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0)
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
        self.buttonchosename.Bind(wx.EVT_BUTTON, self.get_allname)
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
        filename_name = xlrd.open_workbook(filename_allname, encoding_override='cp936')
        namelist = {}
        data_display = {}
        # ������Ա����
        for item_chu in listDivisionName:
            namelist["%s" % item_chu] = []
        for item_chu in listDivisionName:
            handler_sheet = filename_name.sheet_by_name(item_chu)
            total_rows = handler_sheet.nrows
            data_display["%s" % item_chu] = {}
            for index_name in range(0, total_rows):
                name_people = handler_sheet.cell(index_name, 0).value.replace(u' ', u'').strip()
                if name_people not in namelist["%s" % item_chu]:
                    namelist["%s" % item_chu].append(name_people)
                name_people = handler_sheet.cell(index_name, 0).value.replace(u' ', u'').strip()
                data_display["%s" % item_chu]["%s" % name_people] = {}
                data_display["%s" % item_chu]["%s" % name_people]['�����ύ����'.decode('gbk')] = 0
                data_display["%s" % item_chu]["%s" % name_people]['������������'.decode('gbk')] = 0
                data_display["%s" % item_chu]["%s" % name_people]['ʵ�������ύ����'.decode('gbk')] = 0
                data_display["%s" % item_chu]["%s" % name_people]['ʵ��������������'.decode('gbk')] = 0

        # ��������ͳ�����ݣ���ѯ��ͳ��׫дͨ����ר��
        # filename_input = filename_original_zonglan
        file_name = xlrd.open_workbook(filename_original_zonglan, encoding_override='cp936')
        sheet_filter_one = file_name.sheet_by_index(0)
        total_rows_one = sheet_filter_one.nrows
        sheet_filter_two = file_name.sheet_by_index(1)
        total_rows_two = sheet_filter_two.nrows

        list_status = ["׫дͨ��".decode('gbk')]
        # username = ''
        for item_1 in range(1, total_rows_one):
            username = sheet_filter_one.cell(item_1, 6).value
            type_invention = sheet_filter_one.cell(item_1, 4).value.replace(u' ', u'').split(",")[0].strip()
            # shouli_or_not = sheet_filter_one.cell(item_1, 8).value
            status = sheet_filter_one.cell(item_1, 0).value.strip()
            for item_chu in listDivisionName:
                if username in namelist["%s" % item_chu]:
                    # if shouli_or_not != 'None' or status in list_status:
                    if status in list_status:
                        if type_invention == '����'.decode('gbk'):
                            data_display["%s" % item_chu]['%s' % username]['������������'.decode('gbk')] += 1
                        if type_invention == '����'.decode('gbk'):
                            data_display["%s" % item_chu]['%s' % username]['ʵ��������������'.decode('gbk')] += 1
                    if type_invention == '����'.decode('gbk'):
                        data_display["%s" % item_chu]['%s' % username]['�����ύ����'.decode('gbk')] += 1
                    if type_invention == '����'.decode('gbk'):
                        data_display["%s" % item_chu]['%s' % username]['ʵ�������ύ����'.decode('gbk')] += 1

        for item_2 in range(1, total_rows_two):
            username = sheet_filter_two.cell(item_2, 1).value
            type_invention = sheet_filter_two.cell(item_2, 6).value.replace(u' ', u'').strip()
            for item_chu in listDivisionName:
                if username in namelist["%s" % item_chu]:
                    if type_invention == '����'.decode('gbk'):
                        data_display["%s" % item_chu]['%s' % username]['������������'.decode('gbk')] += 1
                    if type_invention == 'ʵ������'.decode('gbk'):
                        data_display["%s" % item_chu]['%s' % username]['ʵ��������������'.decode('gbk')] += 1

        # ��������ͳ�����ݣ���ѯ��ͳ��׫дͨ����ר��
        file_name_shouli = xlrd.open_workbook(filename_original_shouli, encoding_override='cp936')
        sheet_filter_shouli = file_name_shouli.sheet_by_index(0)
        total_rows_shouli = sheet_filter_shouli.nrows
        for item_shouli in range(1, total_rows_shouli):
            username_shouli = sheet_filter_shouli.cell(item_shouli, 4).value.strip()
            type_invention = sheet_filter_shouli.cell(item_shouli, 1).value.split(",")[0].strip()
            for item_chu in listDivisionName:
                if username_shouli in namelist["%s" % item_chu]:
                    if type_invention == '����'.decode('gbk'):
                        data_display["%s" % item_chu]['%s' % username_shouli]['������������'.decode('gbk')] += 1
                    if type_invention == '����'.decode('gbk'):
                        data_display["%s" % item_chu]['%s' % username_shouli]['ʵ��������������'.decode('gbk')] += 1

        # write output excel
        timestamp = time.strftime('%Y%m%d', time.localtime())
        filename_display = "������֤������ר��������ͳ��-%s.xlsx".decode('gbk') % timestamp
        filename_final = os.path.join(dir_filename_display, filename_display)
        workbook_to_write = xlsxwriter.Workbook(filename_final)
        formatone = workbook_to_write.add_format()
        formatone.set_border(1)
        for item_sheet_one in listDivisionName:
            workbook_to_write.add_worksheet(item_sheet_one)
        for item_chu in listDivisionName:
            sheet_now = workbook_to_write.get_worksheet_by_name(item_chu)
            sheet_now.set_column('B:E', 15)
            for i in range(1, len(listTitle) + 1):
                sheet_now.write(0, i, listTitle[i - 1], formatone)

            for index, item_sheet_one in enumerate(namelist["%s" % item_chu]):
                line_count = index + 1
                sheet_now.write(line_count, 0, item_sheet_one, formatone)
                sheet_now.write(line_count, 1,
                                data_display["%s" % item_chu]['%s' % item_sheet_one]['�����ύ����'.decode('gbk')], formatone)
                sheet_now.write(line_count, 2,
                                data_display["%s" % item_chu]['%s' % item_sheet_one]['������������'.decode('gbk')], formatone)
                sheet_now.write(line_count, 3,
                                data_display["%s" % item_chu]['%s' % item_sheet_one]['ʵ�������ύ����'.decode('gbk')],
                                formatone)
                sheet_now.write(line_count, 4,
                                data_display["%s" % item_chu]['%s' % item_sheet_one]['ʵ��������������'.decode('gbk')],
                                formatone)

        print time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
        diag_finish = wx.MessageDialog(None, '�������Ѿ����浽�ļ�%s��%s����.����������������������˳���ť�˳�����'.decode('gbk') % (
            dir_filename_display, filename_display), '��ʾ'.decode('gbk'), wx.OK | wx.ICON_INFORMATION | wx.STAY_ON_TOP)
        diag_finish.ShowModal()
        workbook_to_write.close()


if __name__ == '__main__':
    app = wx.App()
    frame = InventionFilterAll(None)
    frame.Show()
    app.MainLoop()
