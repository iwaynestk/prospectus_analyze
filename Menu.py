# Author: Wayne Yin
# This is the main menu of the whole program. 

import os
import sys

from read_excel import read_excel
from read_pdf import read_pdf
from read_web import read_web

_simple_dir = os.path.join(sys.path[0], 'simplify_pdf')
_write_path = os.path.join(sys.path[0], 'write_excel')
_original_path = os.path.join(sys.path[0], 'original_excel')
_original_dir = os.path.join(sys.path[0], 'original_pdf')

def create_folder():
    create_s = False
    create_wp = False
    create_op = False
    create_od = False

    if not os.path.exists(_simple_dir):
        os.makedirs(_simple_dir)
        create_s = True
    if not os.path.exists(_write_path):
        os.makedirs(_write_path)
        create_wp = True
    if not os.path.exists(_original_path):
        os.makedirs(_original_path)
        create_op = True
    if not os.path.exists(_original_dir):
        os.makedirs(_original_dir)
        create_od = True
    if create_od:
        print("\n已创建original_pdf(需要解析的pdf请放在这里）\n")
    if create_op:
        print("\n已创建original_excel(pdfree解析后下载的excel请放在这里）\n")
    if create_s:
        print("\n已创建simplify_pdf(拆分后的pdf文件夹）\n")
    if create_wp:
        print("\n已创建write_excel(写入的excel文件夹）\n")

class Menu(object):
    def __init__(self):
        self.read_pdf = read_pdf()
        self.read_excel = read_excel()
        self.read_web = read_web()
        self.create_folder = create_folder
        self.choices = {
            "1": self.create_folder,
            "2": self.read_pdf.read_folder,
            "3": self.read_excel.read_folder,
            "4": self.read_web.read_folder,
            "9": self.display_help,
            "0": self.quit
        }

    def display_help(self):
        print("""\n 智能解析招股书脚本V 1.0.0使用帮助：
        
        第一步，请输入1以创建目录。
        第二步，请把待解析的PDF文档放入original_pdf/目录下
        第三步，请输入2开始解析并等待解析完成
        第四步，把simplify_pdf/上传到http://www.pdfree.cn/dashboard解析，并把下载后的excel放入original_excel/目录下
        第五步，输入3开始读取excel数据
        第六步，输入4开始读取证监会网站数据
        Attention:1~6步必须严格按照顺序执行
        
        """)

    def display_menu(self):
        print("""
智能解析招股书 V1.0.0 (BETA) - 资道研究所内部使用:
1. 创建目录
2. 开始解析
3. 从表格读数据
4. 从证监会网站读数据
9. 使用帮助
0. 退出
""")

    def run(self):
        while True:
            self.display_menu()
            try:
                choice = input("Enter an option: ")
            except Exception as e:
                print("Please input a valid option!")
                continue

            choice = str(choice).strip()
            action = self.choices.get(choice)
            if action:
                action()
            else:
                print("{0} is not a valid choice".format(choice))

    def quit(self):
        print("\nThank you for using this script!\n")
        sys.exit(0)


if __name__ == '__main__':
    Menu().run()
