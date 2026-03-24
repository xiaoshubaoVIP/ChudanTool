import os
import sys
from pathlib import Path

import PyQt5
import PyQt5.QtCore as QtCore
from PyQt5.QtWidgets import QHBoxLayout, QVBoxLayout, QWidget, QTextEdit, QMainWindow, QApplication, QAction, \
    QDesktopWidget, QLabel, QLineEdit, QPushButton
import pandas as pd
import openpyxl
from openpyxl import load_workbook

print("pyqt5:v"+QtCore.QT_VERSION_STR)
print(QtCore.QT_VERSION_STR)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.error = '<font color="red">{}</font>'
        self.warning = '<font color="orange">{}</font>'
        self.valid = '<font color="green">{}</font>'

        #设置菜单栏
        self.ex = None
        self.set_bar()

        #文本框设定
        self.path = r"C:\Users\claybox\Desktop\PyTest\2026年1月"
        self.line_edit_path = QLineEdit(self.path)
        self.line_edit_path.setFixedHeight(40)

        #开始按键
        self.push_button_start = QPushButton('开始')
        self.push_button_start.setStyleSheet("background-color: rgb(255,255,255); color: black;")
        self.push_button_start.setFixedSize(100, 40)
        self.push_button_start.clicked.connect(self.start_button)

        #放置一个central_widget
        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)

        #设置水平布局及标题
        main_layout = QVBoxLayout()
        layout_1 = QHBoxLayout()
        layout_2 = QHBoxLayout()

        #实例化路径输入框并加入布局
        layout_1.addWidget(QLabel("输入表格所在目录:"))
        layout_1.addWidget(self.line_edit_path)
        layout_1.addWidget(self.push_button_start)

        #实例化textEdit并加入布局
        self.text_edit = QTextEdit()
        layout_2.addWidget(self.text_edit)

        #实例化plot并加入布局

        #设置窗体的宽高和标题
        self.resize(1680,960)
        self.setWindowTitle("ChuChu小助手 V1.0")

        #将左侧和右侧布局添加到主水平布局中
        main_layout.addLayout(layout_1)
        main_layout.addLayout(layout_2)
        central_widget.setLayout(main_layout)
        #设置窗口的主布局
        self.setCentralWidget(central_widget)

        #调用居中函数
        self.center()


        # 设置信息
        temp_path = QtCore.QDir.currentPath()
        temp_path = QtCore.QDir(temp_path)
        self.set_path = temp_path.absolutePath()+'/setting/'
        if not os.path.isdir(self.set_path):
            os.mkdir(self.set_path)
            print("set文件不存在")
            self.text_edit.append(self.error.format("请先在setting目录下添加设置文件"))
        else:
            set_file = Path(self.set_path + 'setting.xlsx')
            if set_file.is_file():
                try:
                    self.set_data = pd.read_excel(set_file, dtype=str)  # 以字符串形式打开并读取excel表格
                    print(self.set_data)
                    print('set文件存在')
                    self.text_edit.append(self.valid.format("设置文件正常"))
                except FileNotFoundError as e:
                    print(f"设置文件打开失败: {e}")
                    self.text_edit.append(self.error.format("设置文件打开失败"))
            else:
                print("set文件不存在")
                self.text_edit.append(self.error.format("设置文件不存在"))

    def set_bar(self):
        # 实例化主窗口的QMenuBar对象
        bar = self.menuBar()

        # 向菜单栏中添加“文件”
        file = bar.addMenu('文件')
        # 向QMenu小控件中添加按钮，子菜单
        file.addAction('打开文件')
        # 定义响应小控件按钮，并设置快捷键关联到操作按钮，添加到父菜单下
        save = QAction('保存文件', self)
        save.setShortcut('Ctrl+S')
        file.addAction(save)
        # 父菜单下
        quit_ = QAction('退出', self)
        file.addAction(quit_)

        # 向菜单栏中添加“设置”
        menu_set = bar.addMenu('设置')
        menu_set.addAction('配置一')
        menu_set.addAction('配置二')

    def center(self):
        # 获取屏幕的大小
        screen = QDesktopWidget().screenGeometry()
        # 获取窗口的大小
        size = self.geometry()
        # 将窗口移动到屏幕中央
        self.move(int((screen.width() - size.width()) / 2), int((screen.height() - size.height()) / 2))

    @staticmethod
    def company_process(self, dir_name, file_path):
        # print(file_path)

        file_name = os.path.basename(file_path)
        file = Path(file_path)

        if file.is_file():
            data = pd.DataFrame(self.set_data)
            company_data = data[data['公司'] == str(dir_name)]
            print(company_data)
            # print(company_data[['公司', '文件名称摘要', '表格名称', '源数据地址']].head(5))

            # match dir_name:
            #     case '国君':
            #         if "STW895" in file_name:
            #             print('国君:', file_name)
            #             self.text_edit.append('国君:'+file_name)
            #             try:
            #                 wb = load_workbook(file)
            #                 sheet = wb['营改增税负分析测算明细表']
            #                 value1 = sheet['F12'].value     #G20
            #                 value2 = sheet['L12'].value     #I20
            #                 value3 = sheet['G11'].value     #G21
            #                 print(value1, value2, value3)
            #                 self.text_edit.append(str(value1) +','+ str(value2)+',' + str(value3))
            #             except FileNotFoundError as e:
            #                 print(f"Error: {e}")
            #     case '招商':
            #         if "全额申报表" in file_name:
            #             print('招商:', file_name)
            #             self.text_edit.append('招商:'+file_name)
            #             try:
            #                 wb = load_workbook(file)
            #                 sheet = wb['表1-应税申报表']
            #                 value1 = sheet['G15'].value     #G4
            #                 value2 = sheet['I15'].value     #I4
            #                 value3 = sheet['F15'].value     #G5
            #                 print(value1, value2, value3)
            #                 self.text_edit.append(str(value1) +','+ str(value2)+',' + str(value3))
            #             except FileNotFoundError as e:
            #                 print(f"Error: {e}")
            #
            #     case _:
            #         print(dir_name+'无需处理')


    #遍历当前目录
    def list_directory(self):
        for entry in os.listdir(self.path):
            full_path = os.path.join(self.path, entry)

            if os.path.isdir(full_path):                    #如果是目录
                dir_name = os.path.basename(full_path)
                print("目录:"+dir_name)

                data = pd.DataFrame(self.set_data)          #根据目录过滤公司
                company_data = data[data['公司']==str(dir_name)]

                for index in os.listdir(full_path):         #遍历该目录下文件
                    file_path = os.path.join(full_path, index)
                    file_name = os.path.basename(file_path)
                    file = Path(file_path)
                    print("文件:" + file_name)

                    if file.is_file():
                        for i, row in company_data.iterrows():
                            if row['文件名称摘要'] in file_name:
                                print(row['源数据地址'], row['目标地址'])
                                try:
                                    wb = load_workbook(file)
                                    print("open file")
                                    if str(row['表格名称'] in wb.sheetnames):
                                        sheet = wb[str(row['表格名称'])]
                                        src_value = sheet[str(row['源数据地址'])].value     #G20
                                        print(src_value)
                                        self.text_edit.append(str(src_value))
                                    else:
                                        print('源数据文件表格sheet错误')
                                        self.text_edit.append(self.error.format("源数据文件表格sheet错误"))
                                except FileNotFoundError as e:
                                    print(f"Error: {e}")

    def start_button(self):
        print("开始")
        temp_path = QtCore.QDir(self.line_edit_path.text())
        path = temp_path.absolutePath()
        print('目录:'+path)
        self.text_edit.clear()
        self.text_edit.append('目录:'+path)
        if not os.path.isdir(path):
            print("错误")
            self.text_edit.append(self.error.format("错误"))
        else:
            print("正确")
            self.text_edit.append(self.valid.format("目录正常"))
            self.path = path
            self.list_directory()

if __name__ == '__main__':
    # 每一个pyqt程序中都需要有一个QApplication对象，sys.argv是一个命令行参数列表
    app = QApplication(sys.argv)
    # 实例化主窗口并显示
    form = MainWindow()
    form.show()

    # 进入程序的主循环，遇到退出情况，终止程序
    sys.exit(app.exec_())