import threading
import time
from urllib.parse import urljoin

import pandas as pd
from lxml import etree
from selenium import webdriver
from lxml import html

import requests
from PyQt5 import QtCore
from PyQt5.QtCore import QObject
from PyQt5.QtWidgets import QPushButton, QLineEdit, QTextEdit, QVBoxLayout, QHBoxLayout, QWidget, QLabel
from bs4 import BeautifulSoup


class WebCrawler(QWidget):
    def __init__(self):
        super(QWidget, self).__init__()

        self.error = '<font color="red">{}</font>'
        self.warning = '<font color="orange">{}</font>'
        self.valid = '<font color="green">{}</font>'

        #请求状态
        self.request_states = False

        #URL设定
        self.label_url = QLabel("输入URL")
        self.label_url.setStyleSheet("background-color: rgb(255,255,255); color: black;")
        self.label_url.setFixedHeight(40)
        self.line_edit_path = QLineEdit(str('https://gs.amac.org.cn/'))
        self.line_edit_path.setFixedHeight(40)
        self.line_edit_path.setStyleSheet("QLineEdit { background-color: white; }")

        #关键词设定
        self.label_keyword = QLabel("输入搜索关键词")
        self.label_keyword.setStyleSheet("background-color: rgb(255,255,255); color: black;")
        self.label_keyword.setFixedHeight(40)
        self.search_keyword = QLineEdit('进化论')
        self.search_keyword.setFixedSize(120, 40)
        self.search_keyword.setStyleSheet("QLineEdit { background-color: white; }")

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
        stack_layout_1.addWidget(self.label_url)
        stack_layout_1.addWidget(self.line_edit_path)
        stack_layout_1.addWidget(self.label_keyword)
        stack_layout_1.addWidget(self.search_keyword)
        stack_layout_1.addWidget(self.push_button_start)
        stack_layout_2.addWidget(self.text_edit)
        stack_main_layout.addLayout(stack_layout_1)
        stack_main_layout.addLayout(stack_layout_2)
        self.setLayout(stack_main_layout)

        #创建timer
        self.timer = None

    def request_fun(self):
        # headers = {
        #     "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/127.0.0.0 Safari/537.36 Edg/127.0.0.0"
        # }
        # web = self.line_edit_path.text()
        # self.timer = threading.Timer(3, self.repeat_fun, args=(headers, web))
        # self.timer.start()

        self.timer = threading.Timer(3, self.js_load_fun)
        self.timer.start()


    def js_load_fun(self):
        #目标网站
        api_url = str(self.line_edit_path.text()) + '/amac-infodisc/api/pof/fund?&page=0&size=20'

        print(api_url)
        payload = {
            'keyword': str(self.search_keyword.text()),
        }

        headers = {
            # "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/127.0.0.0 Safari/537.36 Edg/127.0.0.0"
            # 'Connection': 'keep-alive',
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/146.0.0.0 Safari/537.36 Edg/146.0.0.0",
            'cookie':'Hm_lvt_a0d0f99af80247cfcb96d30732a5c560=1774920635,1774938609'
            # 某些网站可能需要特定的请求头，如 Referer 或 Authorization
        }

        # 3. 发送 GET 请求
        response = requests.post(api_url, json=payload, headers=headers)
        response.raise_for_status()  # 检查请求是否成功

        # 5. 检查请求是否成功
        try:
            response.raise_for_status()  # 如果状态码不是 200，会抛出异常
            print("请求成功！")
            self.text_edit.append(self.valid.format("请求成功"))
            
            # 6. 解析返回的数据 (通常是 JSON 格式)
            data = response.json()

            # 7. 将数据转换为 DataFrame (以 pandas 处理为例)
            # 你需要根据实际的 JSON 结构调整代码
            # 假设数据在 data['results'] 这个键下面
            print(data)  # 打印完整数据以便分析
            if 'content' in data:
                df = pd.DataFrame(data['content'])
                if df.empty:
                    print("没有搜索到相关记录")
                    self.text_edit.append(self.error.format("没有搜索到相关记录"))
                else:
                    # df.drop(['id', 'managerType', 'lastQuarterUpdate'], axis=1, inplace=True)
                    print(df[['fundNo', 'fundName', 'managerName', 'establishDate', 'putOnRecordDate']])

            else:
                print("未找到预期的数据键，请检查 JSON 结构。")
                print(data)  # 打印完整数据以便分析

        except requests.exceptions.HTTPError as e:
            print(f"请求失败: {e}")

    def repeat_fun(self, headers, url):
        response = requests.get(url, headers=headers)

        response.encoding = "utf-8" # 根据网页实际编码设置
        
        # 确保请求成功
        if response.status_code == 200:
            url = response.text
            print(url)
            # ---------------------------------------------------------
            # tree = html.fromstring(url)
            #
            # # 3. 使用XPath提取表格数据
            # # 直接定位所有tr
            # rows = tree.xpath('//table[@id="fundlist"]//thead//tr')
            #
            # data = []
            # for row in rows:
            #     # 提取当前行中所有的td和th单元格文本
            #     cells = row.xpath('.//td/text() | .//th/text()')
            #     # 清洗数据，去除空白字符
            #     cells = [cell.strip() for cell in cells if cell.strip()]
            #     if cells:
            #         data.append(cells)
            # print(data)

            # ---------------------------------------------------------
            # # 2. 使用BeautifulSoup解析HTML
            # soup = BeautifulSoup(url, 'html.parser')
            #
            # # 3. 定位目标表格 (可以根据id, class等属性精确定位)
            # table = soup.find('table', id='fundlist')
            # # 4. 提取所有行 (tr)
            # rows = table.find_all('tr')
            # # 5. 遍历每一行，提取单元格数据
            # data = []
            # for row in rows:
            #     # 提取当前行中所有的td和th单元格
            #     cells = row.find_all(['td', 'th'])
            #     # 获取单元格的纯文本内容，并去除首尾空格
            #     row_data = [cell.get_text(strip=True) for cell in cells]
            #     if row_data:  # 过滤掉空行
            #         data.append(row_data)
            # # 此时 data 是一个二维列表，包含了整个表格的数据
            # print(data)


            #---------------------------------------------------------
            # 使用BeautifulSoup解析HTML内容
            # soup = BeautifulSoup(url, "html.parser")
            # title = soup.find('title')
            #
            # print(request)
            # print(str(title))

            # ---------------------------------------------------------
            #通过定位 <thead> 下的所有 <th> 标签来获取全部表头，并将每个表头的文本添加到 headers 列表中
            # headers = []
            # header_row = table.find('thead').find_all('th')
            # for th in header_row:
            #     headers.append(th.text.strip())
            # print('header:'+str(headers))

            # ---------------------------------------------------------
            #使用 find_all() 方法提取表格中的每一行并将数据添加到 data 列表
            # data = []
            # Loop through each row in the table (skipping the header row)
            # for tr in table.find_all('tr')[1:]:
            #     row = []
            #     for td in tr.find_all('td'):
            #         cell_data = td.text.strip()
            #         row.append(cell_data)
            #         print(row)
            #     data.append(row)
            #
            #将data转换成dateFrame格式
            # df = pd.DataFrame(data, columns=headers)
            # print(df.shape)

            # ---------------------------------------------------------
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

            # ---------------------------------------------------------
            # tbody = soup1.select('tbody')
            # content = []
            # for row in tbody[0].find_all('tr'):
            #     cols = row.find_all('td')
            #     content.append([col.get_text() for col in cols])
            # print(content)

            # ---------------------------------------------------------
            #找出内部URL
            # links = []
            # for a in soup.find_all("a"):
            #     href = a.get("href")
            #     if href:
            #         links.append(href)
            #
            # for link in links:
            #     print(link)

            # ---------------------------------------------------------
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