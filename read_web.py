# encoding: utf-8
import os

import xlrd
from selenium import webdriver
from bs4 import BeautifulSoup
import re
import sys

from xlutils.copy import copy

digit = {'一': 1, '二': 2, '三': 3, '四': 4, '五': 5, '六': 6, '七': 7, '八': 8, '九': 9}

_path = os.path.join(sys.path[0], 'write_excel')


def _trans(s):
    num = 0
    if s:
        idx_q, idx_b, idx_s = s.find('千'), s.find('百'), s.find('十')
        if idx_q != -1:
            num += digit[s[idx_q - 1:idx_q]] * 1000
        if idx_b != -1:
            num += digit[s[idx_b - 1:idx_b]] * 100
        if idx_s != -1:
            # 十前忽略一的处理
            num += digit.get(s[idx_s - 1:idx_s], 1) * 10
        if s[-1] in digit:
            num += digit[s[-1]]
    return num


def trans(chn):
    chn = chn.replace('零', '')
    idx_y, idx_w = chn.rfind('亿'), chn.rfind('万')
    if idx_w < idx_y:
        idx_w = -1
    num_y, num_w = 100000000, 10000
    if idx_y != -1 and idx_w != -1:
        return trans(chn[:idx_y]) * num_y + _trans(chn[idx_y + 1:idx_w]) * num_w + _trans(chn[idx_w + 1:])
    elif idx_y != -1:
        return trans(chn[:idx_y]) * num_y + _trans(chn[idx_y + 1:])
    elif idx_w != -1:
        return _trans(chn[:idx_w]) * num_w + _trans(chn[idx_w + 1:])
    return _trans(chn)


class read_web(object):

    def startScrapingNames(self, driver, keyWord, wb, path):
        keyWord = keyWord.strip()
        print(keyWord)
        targetUrl = ''
        driver.find_element_by_name("schword").clear()
        driver.find_element_by_name("schword").send_keys(keyWord)
        driver.find_element_by_class_name("so_btn").click()
        html = driver.page_source
        bsobj = BeautifulSoup(html, "html.parser")
        shsj = []
        for result in bsobj.findAll('a'):
            if str(result.text).endswith("工作会议公告"):
                targetUrl = 'http://www.csrc.gov.cn/' + result.attrs['href']
            if re.search(r'[\w\W]*' + keyWord, result.text) and result.text.endswith('报送）'):
                _mat = re.search(r'[\d]{4}年[\d]{1,2}月[\d]{1,2}', result.text)
                shsj.append(_mat.group())
        time = self.find_time(shsj)
        if time:
            print(time)
            ws = wb.get_sheet(1)
            ws.write(2,1,time)
            wb.save(path)
        if targetUrl is '':
            print("没有找到！")
        else:
            print(targetUrl)
            _print = False
            _key_word = False
            driver.get(targetUrl)
            html = driver.page_source
            bsobj = BeautifulSoup(html, "html.parser")
            names = []
            _year = ''

            for result in bsobj.findAll('td'):
                result = result.text.replace('\n', '').strip()
                if result.startswith('发文日期'):
                    print(result.split(':')[1])
                    ws = wb.get_sheet(1)
                    ws.write(3, 1, result.split(':')[1])
                    wb.save(path)

            for link in reversed(bsobj.findAll("p")):
                _mat = re.match(r'^.\、参会发审委委员+', str(link.text).replace(' ', '').strip())
                _mat_1 = re.match(r'[\w\W]*' + keyWord, str(link.text).replace(' ', '').strip())
                if _mat:
                    _print = True
                else:
                    _print = False
                if _mat_1:
                    _key_word = True
                if _key_word and _mat:
                    break
                if _key_word:
                    names += self.splitList(str(link.text).replace(' ', '').strip().split('　　　　'))
            for result in bsobj.findAll('strong'):
                _print = False
                for i in range(0, len(result.text)):
                    if result.text[i] == '年':
                        break
                    if _print:
                        _year += result.text[i]
                    if result.text[i] == '第':
                        _print = True
            print(_year)
            print(names)
            self.writenames(_year, names, wb, path)

    def writenames(self, _year, names, wb, path):
        ws = wb.get_sheet(12)
        year = _year[-4:]
        session = _year[:2]
        ws.write(1, 1, trans(session))
        ws.write(1, 2, year)
        for i in range(0, len(names)):
            ws.write(3, 1 + i, names[i])
        wb.save(path)

    def startScrapingQuestion(self, driver, keyWord, wb, path):
        keyWord = keyWord.strip()
        targetUrl = ''
        html = driver.page_source
        bsobj = BeautifulSoup(html, "html.parser")
        for result in bsobj.findAll('a'):
            if str(result.text).endswith("审核结果公告"):
                targetUrl = 'http://www.csrc.gov.cn/' + result.attrs['href']
                break
        if targetUrl is '':
            print("没有找到！")
        else:
            print(targetUrl)
            driver.get(targetUrl)
            html = driver.page_source
            bsobj = BeautifulSoup(html, "html.parser")
            isStartMatch = False
            questions = []
            for result in bsobj.findAll('p'):
                mat = re.match(r'^[\W]*[0-9]+[\w\W]*', result.text.replace('　　', '').strip())
                if isStartMatch and mat is not None:
                    questions.append(result.text.replace(u'\u3000', u''))
                else:
                    isStartMatch = False

                if re.match(r'[\w\W]*' + keyWord, str(result.text)) is not None:
                    isStartMatch = True
            print(questions)
            self.writequestions(questions, wb, path)

    def writequestions(self, questions, wb, path):
        ws = wb.get_sheet(10)
        for i in range(0, len(questions)):
            if len(questions[i].split('：')) > 1:
                title = questions[i].split('：')[0]
                question = questions[i].split('：')[1]
                ws.write(1 + i, 0, title)
                ws.write(1 + i, 1, question)
            else:
                ws.write(1 + i, 1, questions[i])
        wb.save(path)

    def splitList(self, row):
        name = []
        for i in row:
            print(i)
            if str(i).__contains__('公司') or str(i).__contains__('发行人'):
                break
            else:
                name.append(i.replace(u'\u3000', u''))
                print(i.replace(u'\u3000', u''))
        return name

    def startScraping(self, keyWord, wb, path):
        driver = webdriver.PhantomJS()
        driver.get("http://www.csrc.gov.cn/pub/newsite/")
        if keyWord == '':
            return
        self.startScrapingNames(driver, keyWord, wb, path)
        driver.back()
        self.startScrapingQuestion(driver, keyWord, wb, path)

    def read_excel(self, path, filename):
        print("检索 %s" % filename)
        excel = xlrd.open_workbook(path)
        table = excel.sheet_by_index(0)
        row = table.row_values(0)
        wb = copy(excel)
        if row[1] != '':
            print(row[1])
            self.startScraping(row[1], wb, path)

    def read_folder(self, path=_path):
        if not os.path.exists(path):
            os.makedirs(path)
        files = os.listdir(path)
        for file in files:
            if file.endswith('.xls'):
                self.read_excel(os.path.join(path, file), file)

    def find_time(self,shsj):
        if len(shsj) == 0:
            return None
        _min = shsj[0]
        if len(shsj) >= 2:
            for i in range(len(shsj)):
                if i == 0:
                    continue
                if int(shsj[i].replace('年', '').replace('月', '')) < int(_min.replace('年', '').replace('月', '')):
                    _min = shsj[i]
        return _min
