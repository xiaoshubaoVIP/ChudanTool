import threading
import time
from urllib.parse import urljoin

import pandas as pd
from selenium import webdriver

import requests
from PyQt5 import QtCore
from PyQt5.QtCore import QObject
from PyQt5.QtWidgets import QPushButton, QLineEdit, QTextEdit, QVBoxLayout, QHBoxLayout, QWidget
from bs4 import BeautifulSoup


class WebCrawler(QWidget):
    def __init__(self):
        super(QWidget, self).__init__()
        #请求状态
        self.request_states = False

        #实例化按键
        self.btn = QPushButton("输入URL")
        self.btn.setStyleSheet("background-color: rgb(255,255,255); color: black;")
        self.btn.setFixedHeight(40)
        # self.btn.clicked.connect(self.get_dir)

        #文本框设定
        self.line_edit_path = QLineEdit(str('https://gs.amac.org.cn/amac-infodisc/res/pof/fund/index.html'))
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

        #创建timer
        self.timer = None

    def request_fun(self):
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/127.0.0.0 Safari/537.36 Edg/127.0.0.0"
        }
        web = self.line_edit_path.text()
        self.timer = threading.Timer(3, self.repeat_fun, args=(headers, web))
        self.timer.start()

    def chrome_fun(self):
        #目标网站
        url = self.line_edit_path.text()

        # 创建浏览器驱动对象
        driver_path = r'/Users/weaabduljamac/Downloads/chromedriver'
        driver = webdriver.Chrome(driver_path)

        # 设置代理服务器参数
        proxy_host = "www.16yun.cn"
        proxy_port = "3111"
        proxy_user = "16YUN"
        proxy_pass = "16IP"
        options = webdriver.ChromeOptions()
        options.add_argument(f"--proxy-server=http://{proxy_user}:{proxy_pass}@{proxy_host}:{proxy_port}")

        # 打开目标网站
        driver.get(url)




    def repeat_fun(self, headers, request):
        response = requests.get(request, headers=headers)

        response.encoding = "utf-8" # 根据网页实际编码设置
        
        # 确保请求成功
        if response.status_code == 200:
            html = response.text

            # 使用BeautifulSoup解析HTML内容
            soup = BeautifulSoup(html, "html.parser")
            title = soup.find('title')

            print(request)
            print(str(title))

            table = soup.find('table', attrs={'id': 'fundlist'})
            print(str(table))

            #通过定位 <thead> 下的所有 <th> 标签来获取全部表头，并将每个表头的文本添加到 headers 列表中
            headers = []
            header_row = table.find('thead').find_all('th')
            for th in header_row:
                headers.append(th.text.strip())
            print('header:'+str(headers))

            #使用 find_all() 方法提取表格中的每一行并将数据添加到 data 列表
            data = []
            # Loop through each row in the table (skipping the header row)
            for tr in table.find_all('tr')[1:]:
                row = []
                for td in tr.find_all('td'):
                    cell_data = td.text.strip()
                    row.append(cell_data)
                    print(row)
                data.append(row)

            #将data转换成dateFrame格式
            df = pd.DataFrame(data, columns=headers)
            print(df.shape)

            #提取css文件
            # print("---css_url---")
            # for link in soup.find_all('link', rel='stylesheet'):
            #     css_url = link['href']
            #     print(css_url)
            #
            #提取js文件
            # print("+++js_url+++")
            # for script in soup.find_all('script', src=True):
            #     js_url = script['src']
            #     print(js_url)


            # tbody = soup1.select('tbody')
            # content = []
            # for row in tbody[0].find_all('tr'):
            #     cols = row.find_all('td')
            #     content.append([col.get_text() for col in cols])
            # print(content)


            #找出内部URL
            # links = []
            # for a in soup.find_all("a"):
            #     href = a.get("href")
            #     if href:
            #         links.append(href)
            #
            # for link in links:
            #     print(link)

            #百度热搜
            # for tt in soup.find_all('ul', class_='s-hotsearch-content'):
            #     for ss  in tt.find_all('span', class_='title-content-title'):
            #         print(ss.text)
            #         self.text_edit.append(str(ss.text))
            # self.text_edit.append("--------------------------\n")

        else:
            print("请求失败，状态码：", response.status_code)

        #重复执行
        self.request_fun()

    #复写closeEvent
    def stack_close(self):
        print("关闭timer")
        self.push_button_start.setEnabled(False)
        self.timer.cancel()

    def start_button(self):
        if self.request_states:
            self.request_states = False
            self.push_button_start.setText('请求')
            self.timer.cancel()
        else:
            self.request_states = True
            print('请求开始')
            self.push_button_start.setText('停止')
            self.request_fun()