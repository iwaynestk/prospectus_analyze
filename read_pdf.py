# encoding: utf-8
from pdfminer.pdfparser import PDFParser, PDFDocument
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter, PDFTextExtractionNotAllowed
from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LTTextBoxHorizontal, LAParams
from PyPDF2 import PdfFileReader, PdfFileWriter

import sys
import importlib
import os
import warnings
import xlrd
import re
import shutil

from xlutils.copy import copy

warnings.filterwarnings('ignore')
importlib.reload(sys)

_simple_dir = os.path.join(sys.path[0], 'simplify_pdf')
_path = os.path.join(sys.path[0], 'write_excel')
_original_dir = os.path.join(sys.path[0], 'original_pdf')


class read_pdf(object):

    def read_folder(self, folder_path=_original_dir):
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)
        files = os.listdir(folder_path)
        findPdf = False
        for file in files:
            if file.endswith('pdf'):
                findPdf = True
                self.parse(os.path.join(folder_path, file), file)
        if not findPdf:
            print("指定目录未找到pdf文档")

    def parse(self, path, filename):
        print('----------------------------------------------------------')
        print('查找文档:' + filename)

        writepath = self.copy_excel('s' + filename.rsplit('.', 1)[0])
        wb = copy(xlrd.open_workbook(writepath, formatting_info=True))
        ws = wb.get_sheet(0)

        fp = open(path, 'rb')  # 以二进制读模式打开
        # 用文件对象来创建一个pdf文档分析器
        praser = PDFParser(fp)
        # 创建一个PDF文档
        doc = PDFDocument()
        # 连接分析器 与文档对象
        praser.set_document(doc)
        doc.set_parser(praser)
        # 提供初始化密码
        # 如果没有密码 就创建一个空的字符串
        doc.initialize()

        # 检测文档是否提供txt转换，不提供就忽略
        if not doc.is_extractable:
            raise PDFTextExtractionNotAllowed
        else:
            # 创建PDF 资源管理器 来管理共享资源
            rsrcmgr = PDFResourceManager()
            # 创建一个PDF设备对象
            laparams = LAParams()
            device = PDFPageAggregator(rsrcmgr, laparams=laparams)
            # 创建一个PDF解释器对象
            interpreter = PDFPageInterpreter(rsrcmgr, device)
            # 设置参数
            # 循环遍历列表，每次处理一个page的内容
            flag = False
            count = 0
            _page_finance = 0
            _page_agency = 0
            _page_overview = 0
            _page_money = 0
            _count__read_page = 0
            layouts = []
            _mat_shiyi = False
            for page in doc.get_pages():  # doc.get_pages() 获取page列表
                if _count__read_page == 3:
                    ws = wb.get_sheet(7)
                    self.findPartyConcernedMsg(layouts, writepath, wb, ws)
                    break
                interpreter.process_page(page)
                layout = device.get_result()
                count += 1
                if _count__read_page != 0:
                    _count__read_page += 1
                    layouts.append(layout)
                    continue
                # _read_row = 0

                for x in layout:
                    if isinstance(x, LTTextBoxHorizontal):
                        # _read_row += 1
                        results = x.get_text().replace(" ", "").replace('\n', '').strip()
                        if self.locationYW(results):
                            print(results)
                        # if re.match(r'[\w\W]*第一节[\W]*释义', results):
                        #     _mat_shiyi = True
                        # if not _mat_shiyi and _read_row > 5:
                        #     break
                        # if filename == '晶丰明源.pdf':
                        #     print(results)
                        if re.match('[\w\W]*目录+[\w\W]*', results):
                            break
                        if self.locationOverview(results):
                            print('找到概览，在第 %d 页' % count)
                            _page_overview = count
                            self.findOverviewMsf(layout, writepath, wb, ws)
                            break
                        if self.findPartyConcerned(results):
                            print('找到中介机构,在第 %d 页 ' % count)
                            _count__read_page += 1
                            layouts.append(layout)
                            _page_agency = count
                            self.split_pdf(_page_finance, _page_money, _page_agency, _page_overview, path,
                                           self.create_spdf_dir(filename))
                            break
                        mat = re.search(r'[二|三|四|五]+\、[\w\W]*主要财务数据[\w]*', results)
                        _mat = re.search(r'[\w\W]+\、募集资金[\w\W]*(用途|运用)[\w\W]*', results)
                        if _mat is not None:
                            _page_money = count
                            print("找到募集资金 在第 %d 页" % count)
                        # 募集资金用途:在第三节财务数据的后边。

                        if mat is not None:
                            if self.findCurrentLiabilities(layout):
                                # locationSuccess = True
                                _page_finance = count
                                # split_pdf(count, path, create_spdf_dir(filename))
                                print('找到财务信息,在第 %d 页' % count)
                                break
                            else:
                                flag = True
                                break  # 结束当前页
                        else:
                            if flag:
                                if self.findCurrentLiabilities(layout):
                                    # locationSuccess = True
                                    _page_finance = count
                                    # split_pdf(count, path, create_spdf_dir(filename))
                                    print('找到财务信息,第 %d 页' % count)
                                    break
                                else:
                                    flag = False

    def locationYW(self, results):
        _mat = re.search(r'第[五|六|七]节业务[和|与]技术\.*([\d]{2,3})',results)
        if _mat:
            print("找到主营业务：" + _mat.group(1))
            return True
        else:
            return False

    def findOverviewMsf(self, layout, path, wb, ws):
        time = []
        for x in layout:
            if isinstance(x, LTTextBoxHorizontal):
                results = x.get_text().replace(" ", "").replace('\n', '').strip()
                _mat_name = re.search(r'[\w\W]*(公司|企业|中文)名称', results)
                _mat_money = re.search(r'[\w\W]*注册资本', results)
                # _mat_datetime = re.search(r'[\w\W]*(设立时间|成立日期)', results)
                _mat = re.search(r'[\d]{4}年[\d]{1,2}月[\d]{1,2}[日]?', results)
                _mat_address = re.search(r'[\w\W]*(住所|注册地址)', results)
                if _mat_name:
                    if len(results.split('：')) > 1:
                        print("找到名字：" + results.split('：')[1])
                        ws.write(0, 1, results.split('：')[1])
                        continue
                if _mat_money:
                    if len(results.split('：')) > 1:
                        print("找到注册资本：" + results.split('：')[1])
                        ws.write(6, 1, results.split('：')[1])
                        continue
                if _mat:
                    time.append(_mat.group())
                    continue
                # if _mat_datetime:
                #     if len(results.split('：')) > 1:
                #         print("找到成立日期：" + results.split('：')[1])
                #         ws.write(1, 1, results.split('：')[1])
                #         continue
                if _mat_address:
                    if len(results.split('：')) > 1:
                        print("找到地址：" + results.split('：')[1])
                        ws.write(2, 1, results.split('：')[1])
                        continue

        if len(time) == 1:
            ws.write(1, 1, time[0])
            ws = wb.get_sheet(1)
            ws.write(0, 1, time[0])
        elif len(time) == 0:
            pass
        else:
            ws.write(1, 1, time[0])
            ws = wb.get_sheet(1)
            ws.write(0, 1, time[0])
            ws.write(1, 1, time[1])

        wb.save(path)

    def locationOverview(self, results):
        _mat = re.match(r'^第.节[\w\W]*概览', results)
        if _mat:
            print(results)
            return True
        return False

    def copy_excel(self, filename):
        if not os.path.exists(_path + '/' + filename + '.xls'):
            filepath = shutil.copy(_path + '/IPO.xls', _path + '/' + filename + '.xls')
            return filepath
        else:
            return _path + '/' + filename + '.xls'

    # 检索中介机构的详细信息：保荐人,律所,会计师事务所
    def findPartyConcernedMsg(self, layouts, path, wb, ws):

        _match_sponsor_title = False
        _match_law_title = False
        _match_cpa_title = False

        _match_sponsor_principal = False
        _match_law_principal = False
        _match_cpa_principal = False

        _match_sponsor_manager = False
        _match_law_manager = False
        _match_cpa_manager = False

        for layout in layouts:
            for x in layout:
                if isinstance(x, LTTextBoxHorizontal):
                    results = x.get_text().replace(" ", "").replace('\n', '').strip()
                    print(results)
                    _mat_sponsor = re.match(r'[\w\W]{0,4}保荐+(人|机构)+[\w\W]*', results)
                    _mat_sponsor_add = re.match(r'[\w\W]*收款银行+[\w\W]*', results)
                    _mat_cpa = re.match(r'[\w\W]{0,5}(会计师事务所|公司审计机构)+[\w\W]*', results)
                    _mat_law_office = re.match(r'[\w\W]{0,5}(发行人律师|律师事务所|公司律师)+[\w\W]*', results)
                    _mat_sponsor_principal = re.match(r'法定代表人+', results)
                    _mat_sponsor_manager = re.match(r'保荐代表人+', results)
                    _mat_cpa_principal = re.match(r'[\w\W]*(负责人|法定代表人)+', results)
                    _mat_cpa_manager = re.match(r'(经办|签字)+(注册)*会计师+', results)
                    _mat_law_principal = re.match(r'[\w\W]*负责人+', results)
                    _mat_law_manager = re.match(r'经办律师+', results)
                    if _match_sponsor_title and not _match_sponsor_manager:
                        if _mat_sponsor_principal and len(results.split('：')) > 1:  # 找到保荐法定代表人
                            ws.write(2, 1, results.split('：')[1])
                            print(results[1])
                        if _mat_sponsor_manager:  # 找到保荐代表人
                            _match_sponsor_manager = True
                            if len(results.split('：')) > 1:
                                print(results)
                                if len(results.split('：')[1].split('、')) > 1:
                                    ws.write(3, 1, results.split('：')[1].split('、')[0])
                                    ws.write(4, 1, results.split('：')[1].split('、')[1])
                                else:
                                    ws.write(3, 1, results.split('：')[1])
                    if _match_law_title and not _match_law_manager:
                        if _mat_law_principal and len(results.split('：')) > 1:  # 找到律所负责人
                            ws.write(2, 7, results.split('：')[1])
                            print(results)
                        if _mat_law_manager:  # 找到律所经办律师
                            _match_law_manager = True
                            if len(results.split('：')) > 1:
                                print(results)
                                if len(results.split('：')[1].split('、')) > 1:
                                    ws.write(3, 7, results.split('：')[1].split('、')[0])
                                    ws.write(4, 7, results.split('：')[1].split('、')[1])
                                else:
                                    ws.write(3, 7, results.split('：')[1])
                    if _match_cpa_title and not _match_cpa_manager:
                        if _mat_cpa_principal and len(results.split('：')) > 1:
                            ws.write(2, 4, results.split('：')[1])
                            print(results)
                        if _mat_cpa_manager:
                            _match_cpa_manager = True
                            if len(results.split('：')) > 1:
                                print(results)
                                if len(results.split('：')[1].split('、')) > 1:
                                    ws.write(3, 4, results.split('：')[1].split('、')[0])
                                    ws.write(4, 4, results.split('：')[1].split('、')[1])
                                else:
                                    ws.write(3, 4, results.split('：')[1])
                    if _mat_sponsor and not _mat_sponsor_add:  # 找到保荐机构title
                        sp = results.split('：')
                        _match_sponsor_title = True
                        if len(sp) > 1:
                            ws.write(1, 1, sp[1])
                            print(sp[1])
                    if _mat_cpa:  # 找到会计事务所
                        sp = results.split('：')
                        _match_cpa_title = True
                        if len(sp) > 1:
                            ws.write(1, 4, sp[1])
                            print(sp[1])
                    if _mat_law_office:  # 找到律所
                        _match_law_title = True
                        sp = results.split('：')
                        if len(sp) > 1:
                            ws.write(1, 7, sp[1])
                            print(sp[1])
        wb.save(path)

    # 精确定位中介机构信息在pdf文档中的位置
    def findPartyConcerned(self, results):
        _mat_1 = re.match(r'^[\w\W]*本次[\w\W]*发行[\w\W]*有关[\w]?当事人$', results)
        _mat_2 = re.match(r'^[\w\W]*本次[\w\W]*发行[\w\W]*有关[\w]?机构$', results)
        if _mat_1 is not None or _mat_2 is not None:
            print(results)
            return True
        return False

    # 生成目录
    def create_spdf_dir(self, file_name):
        folder = os.path.exists(_simple_dir)
        if not folder:
            os.makedirs(_simple_dir)
        file_name = 's' + file_name
        path = os.path.join(_simple_dir, file_name)
        file = open(path, 'w')
        file.close()
        return path

    # 拆分pdf
    def split_pdf(self, _page_finance, _page_money, _page_agency, _page_overview, intpath, outputpath):
        pdf_output = PdfFileWriter()
        pdf_input = PdfFileReader(intpath)
        # print(_page_finance)
        # print(_page_money)
        # print(_page_agency)
        if _page_overview != 0:
            pdf_output.addPage(pdf_input.getPage(_page_overview - 1))
        if _page_finance != 0:
            for i in range(3):
                pdf_output.addPage(pdf_input.getPage(i + _page_finance - 1))
        if (_page_money - _page_finance) <= 2:
            if _page_money != (_page_agency - 1):
                pdf_output.addPage(pdf_input.getPage(_page_money))
        else:
            pdf_output.addPage(pdf_input.getPage(_page_money - 1))
            if _page_money != (_page_agency - 1):
                pdf_output.addPage(pdf_input.getPage(_page_money))
        if _page_agency != 0:
            for i in range(2):
                pdf_output.addPage(pdf_input.getPage(i + _page_agency - 1))
        pdf_output.write(open(outputpath, 'wb'))

    # 定位财务表
    def findCurrentLiabilities(self, layout):
        show = False
        for x in layout:
            if isinstance(x, LTTextBoxHorizontal):

                result = x.get_text().replace(" ", "").replace('\n', '').strip()
                if re.match(r'[\w\W]*目录+[\w\W]*', result) is not None:
                    break
                mat = re.search(r'[\w\W]*\（[\w]*\）[\w\W]*资产负债表[\w]*', result)
                if mat is not None:
                    show = True
                    break
        return show
