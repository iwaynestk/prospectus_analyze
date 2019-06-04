# -*- coding: utf-8 -*-
import os
import sys

import xlrd
import re
import shutil
from xlutils.copy import copy

_path = os.path.join(sys.path[0], 'write_excel')

_original_dir = os.path.join(sys.path[0], 'original_excel')


class read_excel(object):

    def location_jzc(self, row):
        for item in row:
            mat1 = re.match(r'^股东权益[\w]*', item)
            mat2 = re.match(r'净资产', item)
            mat3 = re.match(r'所有者权益[\w]*', item)
            if mat1 is not None or mat2 is not None or mat3 is not None:
                return item
            else:
                return None

    def location_yysr(self, row):
        for item in row:
            mat1 = re.match(r'营业[\w]*收入', item)
            if mat1 is not None:
                return item
            else:
                return None

    def location_jlr(self, row):
        for item in row:
            mat1 = re.match(r'^净利润', item)
            if mat1 is not None:
                return item
            else:
                return None

    def location_zzc(self, row):
        for item in row:
            mat1 = re.match(r'^总资产$', item)
            mat2 = re.match(r'资产[\w]*计', item)
            mat3 = re.match(r'资产总额', item)
            if mat1 is not None or mat2 is not None or mat3 is not None:
                return item
            else:
                return None

    def location_zfz(self, row):
        for item in row:
            mat1 = re.match(r'^总负债$', item)
            mat2 = re.match(r'负债[\w]*计', item)
            mat3 = re.match(r'负债总额', item)
            if mat1 is not None or mat2 is not None or mat3 is not None:
                return item
            else:
                return None

    def location_xxjll(self, row):
        for item in row:
            mat1 = re.match(r'经营活动[产生的]*现金流量净额 ', item)
            if mat1 is not None:
                return item
            else:
                return None

    def read_folder(self, path=_original_dir):
        if not os.path.exists(path):
            os.makedirs(path)
        files = os.listdir(path)
        for file in files:
            if file.endswith('.xlsx'):
                self.read_excel(os.path.join(path, file), file)

    def read_excel(self, path, filename):
        print("find %s" % filename)
        excel = xlrd.open_workbook(path)
        writepath = self.copy_excel(filename[:-5])
        ipo = xlrd.open_workbook(writepath, formatting_info=True)
        wb = copy(ipo)
        zzc = []
        zfz = []
        isLocationJzc = False
        _sheet_count = 0
        _sheets = []
        self.findPartyConcernedMsg(excel, writepath, wb)
        ws = wb.get_sheet(0)
        self.findCompanyMsg(excel.sheet_by_index(0), writepath, wb, ws)
        self.findMoney(excel,writepath,wb,ws)
        ws = wb.get_sheet(2)
        for sheet_name in excel.sheet_names():
            table = excel.sheet_by_name(sheet_name)
            nrows = table.nrows
            for rownum in range(0, nrows):
                row = table.row_values(rownum)
                # print(row)
                if self.location_jzc(row):
                    isLocationJzc = True
                    self.write_to_excel(row, 8, writepath, ws, wb)
                    # print(row)
                    # print("找到净资产")
                    # print("-------------------------------------------")
                if self.location_yysr(row):
                    self.write_to_excel(row, 2, writepath, ws, wb)
                    # print(row)
                    # print("找到营业收入")
                    # print("-------------------------------------------")
                if self.location_jlr(row):
                    self.write_to_excel(row, 3, writepath, ws, wb)
                    # print(row)
                    # print("找到净利润")
                    # print("-------------------------------------------")
                if self.location_zzc(row):
                    zzc = self.checkRow(row)[1:]
                    self.write_to_excel(row, 7, writepath, ws, wb)
                    # print(row)
                    # print("找到总资产")
                    # print("-------------------------------------------")
                if self.location_xxjll(row):
                    self.write_to_excel(row, 4, writepath, ws, wb)
                    # print(row)
                    # print("找到现金流量")
                    # print("-------------------------------------------")
                if self.location_zfz(row):
                    zfz = self.checkRow(row)[1:]
                    # print(zfz)
                    # print("找到总负债")
                    # print("-------------------------------------------")
        if not isLocationJzc:
            row = list(map(self.reduceNum, self.str2float(zzc), self.str2float(zfz)))
            row.insert(0, '净资产')
            self.write_to_excel(row, 8, writepath, ws, wb)
            # print(row)
            # print("通过总资产-总负债的方式 找到净资产")
            # print("-------------------------------------------")

    def findMoney(self, excel, path, wb, ws):
        ws = wb.get_sheet(5)
        for sheet_name in excel.sheet_names():
            table = excel.sheet_by_name(sheet_name)
            rows = table.nrows
            _print = False
            __print = False
            _write_count = 0
            for rownum in range(0,rows):
                row = table.row_values(rownum)
                filter(lambda x:x != '',row)
                _mat = re.search(r'[\w\W]*项目名称[\w\W]*',str(row[0]))
                if len(row) > 1:
                    __mat = re.search(r'[\w\W]*项目名称[\w\W]*',str(row[1]))
                if row[0] == 'Made by PDFree(鼎复数据出品)':
                    _print = False
                    __print = False
                    break
                if _mat:
                    _print = True
                    print("找到募集资金（没有序号）：")
                    print(row)
                    print("-------------------------------------")
                    if len(str(row[0]).split(' ',1)) >= 2 and row[0].split(' ',1)[0] != '':
                        if row[0].split(' ',1)[1] != '':
                            print(list(map(lambda x:str(x).split(' ',1)[0],row)))
                            ws.write(1,1,row[1].split(' ',1)[1])
                            ws.write(1,2,row[2].split(' ',1)[1])
                            ws.write(1,3,row[2].split(' ',1)[1])
                            _write_count += 1
                            continue
                if __mat:
                    __print = True
                    print("找到募集资金（有序号）：")
                    print(row)
                    print("-------------------------------------")
                    if len(str(row[0]).split(' ',1)) >= 2 and row[0].split(' ',1)[0] != '':
                        if row[0].split(' ',1)[1] != '':
                            print(list(map(lambda x:str(x).split(' ',1)[0],row)))
                            ws.write(1, 1, row[1].split(' ', 1)[1])
                            ws.write(1, 2, row[2].split(' ', 1)[1])
                            ws.write(1, 3, row[2].split(' ', 1)[1])
                            _write_count += 1
                            continue
                if _print:
                    print(row)
                    _write_count += 1
                    ws.write(_write_count, 1, row[0])
                    ws.write(_write_count, 2, row[1])
                    ws.write(_write_count, 3, row[1])
                if __print:
                    print(row)
                    _write_count += 1
                    ws.write(_write_count, 1, row[1])
                    ws.write(_write_count, 2, row[2])
                    ws.write(_write_count, 3, row[2])
        wb.save(path)

    def findCompanyMsg(self, table, path, wb, ws):
        nrows = table.nrows
        time = []
        for rownum in range(0, nrows):
            row = table.row_values(rownum)
            results = row[0]
            value = ''
            if len(row) > 1:
                value = str(row[1])
            print(value)
            _mat = re.search(r'[\d]{4}年[\d]{1,2}月[\d]{1,2}[日]?', value)
            _mat_name = re.search(r'[\w\W]*(公司|企业|中文)名称', results)
            _mat_money = re.search(r'[\w\W]*注册资本', results)
            # _mat_datetime = re.search(r'[\w\W]*(设立时间|((有限公司)*成立日期))*', results)
            _mat_address = re.search(r'[\w\W]*(住所|注册地址)*', results)
            if _mat_name:
                ws.write(0, 1, row[1])
            if _mat_money:
                ws.write(6, 1, row[1])
            if _mat:
                time.append(row[1])
                # ws.write(1, 1, row[1])
            if _mat_address:
                ws.write(2, 1, row[1])
        if len(time) == 1:
            ws.write(1, 1, time[0])
            ws = wb.get_sheet(1)
            ws.write(0, 1, time[0])
        elif len(time) == 0:
            pass
        else:
            ws.write(1,1,time[0])
            ws = wb.get_sheet(1)
            ws.write(0, 1,time[0])
            ws.write(1,1,time[1])
        wb.save(path)

    # 检索中介机构的详细信息：保荐人,律所,会计师事务所
    def findPartyConcernedMsg(self, excel, path, wb):
        ws = wb.get_sheet(7)
        _match_sponsor_title = False
        _match_law_title = False
        _match_cpa_title = False

        _match_sponsor_principal = False
        _match_law_principal = False
        _match_cpa_principal = False

        _match_sponsor_manager = False
        _match_law_manager = False
        _match_cpa_manager = False

        for sheet_name in reversed(excel.sheet_names()):
            table = excel.sheet_by_name(sheet_name)
            nrows = table.nrows
            for rownum in range(nrows - 1, -1, -1):
                row = table.row_values(rownum)
                results = row[0]
                _mat_sponsor_principal = re.match(r'(法定代表人)+', results)
                _mat_sponsor_manager = re.match(r'[\w\W]*保荐代表人+', results)
                _mat_cpa_principal = re.match(r'[\w\W]*(负责人|法定代表人)+', results)
                _mat_cpa_manager = re.match(r'(经办|签字)+(注册)*会计师+', results)
                _mat_law_principal = re.match(r'(负责人)+', results)
                _mat_law_manager = re.match(r'(经办律师)+', results)
                if _mat_cpa_manager and not _match_cpa_manager:
                    _match_cpa_manager = True  # 找到经办注册会计师
                    if len(row[1].split('、')) > 1:
                        ws.write(3, 4, row[1].split('、')[0])
                        ws.write(4, 4, row[1].split('、')[1])
                    else:
                        ws.write(3, 4, row[1])
                if _match_cpa_manager and not _match_cpa_principal and _mat_cpa_principal:
                    _match_cpa_principal = True  # 找到会计师负责人
                    ws.write(2, 4, row[1])
                if _mat_law_manager and not _match_law_manager:
                    _match_law_manager = True  # 找到经办律师
                    if len(row[1].split('、')) > 1:
                        ws.write(3, 7, row[1].split('、')[0])
                        ws.write(4, 7, row[1].split('、')[1])
                    else:
                        ws.write(3, 7, row[1])
                if _match_law_manager and not _match_law_principal and _mat_law_principal:
                    _match_law_principal = True  # 找到律师负责人
                    ws.write(2, 7, row[1])
                if _mat_sponsor_manager and not _match_sponsor_manager:
                    _match_sponsor_manager = True  # 找到保荐代表人
                    if len(row[1].split('、')) > 1:
                        ws.write(3, 1, row[1].split('、')[0])
                        ws.write(4, 1, row[1].split('、')[1])
                    else:
                        ws.write(3, 1, row[1])
                if _match_sponsor_manager and not _match_sponsor_principal and _mat_sponsor_principal:
                    _match_sponsor_principal = True  # 找到保荐法定代表人
                    ws.write(2, 1, row[1])
        wb.save(path)

    def reduceNum(self, x1, x2):
        return str(format(x1 - x2, '.2f'))

    def str2float(self, row):
        for i in range(0, len(row)):
            row[i] = float(row[i].replace(',', ''))
        return row

    def copy_excel(self, filename):
        if not os.path.exists(_path + '/' + filename + '.xls'):
            filepath = shutil.copy(_path + '/' + 'IPO.xls', _path + '/' + filename + '.xls')
            return filepath
        else:
            return _path + '/' + filename + '.xls'

    def write_to_excel(self, row, startpage, writepath, ws, wb):
        try:
            row = self.checkRow(row)
            # print(row)
            for i in range(0, 3):
                ws.write(startpage, i + 1, row[i + 1])
                wb.save(writepath)
        except IndexError:
            print('表格数据检查有误')

    def checkRow(self, row):
        nRow = []
        isEdit = False
        for i in range(0, len(row)):
            row[i] = str(row[i]).replace(' ', '')
            if i == 0:
                if re.match(r'[\w\W]*[0-9]$', row[i]) is not None:  # 首个单 元格检查  是否是数字结尾，如果是，则分割。
                    for j in range(0, row[i].__len__()):
                        if re.match('[0-9]', row[i][j]):
                            isEdit = True
                            nRow.append(row[i][:j])
                            nRow.append(row[i][j:])
                            break
                else:
                    nRow.append(self.append(i, row))
            else:
                if row[i].count('.') > 1:  # 检查数字中 是否合并了单元格
                    isEdit = True
                    nRow.append(row[i][:row[i].index('.') + 3])
                    temp = row[i][row[i].index('.') + 3:]  # 可以优化为递归算法
                    if temp.count('.') > 1:
                        nRow.append(temp[:temp.index('.') + 3])
                        nRow.append(temp[temp.index('.') + 3:])
                    else:
                        nRow.append(temp)
                else:
                    nRow.append(self.append(i, row))
        if isEdit:
            return nRow
        else:
            return row

    # def splitByDot(row, s, i):
    #     if s.count('.') > 1:
    #         row.append(s[:s.index('.')+3])
    #         return splitByDot(row, s[s.index('.')+3:], i)
    #     else:
    #         row.append(append(i, row))
    #         return row
    '''
    此函数的作用是依次向后找到不为空的那个值，填入list
    '''

    def append(self, i, row):
        if i >= len(row):
            return ''
        if row[i] != '':
            return row[i]
        else:
            if i < len(row):
                return self.append(i + 1, row)
            else:
                return ''
