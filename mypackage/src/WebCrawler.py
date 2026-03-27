import threading
import time
import requests
from PyQt5 import QtCore
from PyQt5.QtCore import QObject
from PyQt5.QtWidgets import QPushButton, QLineEdit, QTextEdit, QVBoxLayout, QHBoxLayout, QWidget
from bs4 import BeautifulSoup


class WebCrawler(QWidget):
    def __init__(self):
        super(QWidget, self).__init__()
        #实例化按键
        self.btn = QPushButton("输入URL")
        self.btn.setStyleSheet("background-color: rgb(255,255,255); color: black;")
        self.btn.setFixedHeight(40)
        # self.btn.clicked.connect(self.get_dir)

        #文本框设定
        self.line_edit_path = QLineEdit(str('https://www.baidu.com/Index.htm'))
        self.line_edit_path.setFixedHeight(40)
        self.line_edit_path.setStyleSheet("QLineEdit { background-color: white; }")

        #开始按键
        self.push_button_start = QPushButton('请求')
        self.push_button_start.setStyleSheet("background-color: rgb(255,255,255); color: black;")
        self.push_button_start.setFixedSize(100, 40)
        self.push_button_start.clicked.connect(self.start_button)

        #实例化textEdit并加入布局
        self.text_edit = QTextEdit()
        self.text_edit.setStyleSheet("QTextEdit { background-color: white; }")

        stack_main_layout = QVBoxLayout()
        stack_layout_1 = QHBoxLayout()
        stack_layout_2 = QHBoxLayout()

        #并加入布局
        stack_layout_1.addWidget(self.btn)
        stack_layout_1.addWidget(self.line_edit_path)
        stack_layout_1.addWidget(self.push_button_start)
        stack_layout_2.addWidget(self.text_edit)
        stack_main_layout.addLayout(stack_layout_1)
        stack_main_layout.addLayout(stack_layout_2)
        self.setLayout(stack_main_layout)


    def request_fun(self):
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/127.0.0.0 Safari/537.36 Edg/127.0.0.0"
        }
        web = self.line_edit_path.text()
        timer1 = threading.Timer(3, self.repeat_fun, args=(headers, web))
        timer1.start()

    def repeat_fun(self, headers, request):
        response = requests.get(request, headers=headers)
        # 确保请求成功
        if response.status_code == 200:
            html = response.text

            # 使用BeautifulSoup解析HTML内容
            soup = BeautifulSoup(html, "html.parser")
            print(request)
            print(soup.find('title'))

            # 查找所有标题（<span>），提取"class"属性为"title"的元素
            # all_titles = soup.findAll("span", attrs={"class": "title"})
            #
            # for title in all_titles:
            #     title_string = title.string
            #     if '/' not in title_string:
            #         print(title_string)
        else:
            print("请求失败，状态码：", response.status_code)

        #重复执行
        self.request_fun()

    def start_button(self):
        print('请求开始')
        self.request_fun()
        # thread1.join()

