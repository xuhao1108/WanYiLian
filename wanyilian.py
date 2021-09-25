#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time : 2021/7/29 11:24
# @Author : YXH
# @Email : 874591940@qq.com
# @desc : ...
import time
import xlrd
import openpyxl

from threading import Thread, Lock

from selenium.webdriver import Chrome, ChromeOptions
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as ec

# 存放商品url
urls = []
# 线程锁
my_lock = Lock()
# 获取商品url的线程，获取商品详细信息的线程
page_thread, details_thread = None, None


class Base(object):
    def __init__(self, config):
        """
        初始化参数，打开浏览器
        :param config: 配置参数
        """
        options = ChromeOptions()
        # 接管已打开的浏览器
        # chrome.exe --remote-debugging-port=9222
        # options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
        # 用户数据位置
        # chrome.exe --user-data-dir="......"
        # options.add_argument(r'--user-data-dir={}'.format(chrome_data_path))
        # 最大化
        # options.add_argument("--start-maximized")
        # 无窗口化
        # options.add_argument('--headless')
        # 禁用GPU加速
        options.add_argument('--disable-gpu')
        # 忽略证书错误
        options.add_argument('--ignore-certificate-errors')
        #
        # options.add_experimental_option('excludeSwitches', ['enable-logging'])
        # 代理
        # options.add_extension(create_proxy_auth_extension('host', 'port', 'user', 'password'))
        # 创建浏览器
        self.chrome = Chrome(options=options)
        self.wait = WebDriverWait(self.chrome, 30)
        self.config = config
        self.max_page = 1

    def get_element_by_xpath(self, pattern):
        """
        通过xpath获取元素
        :param pattern: 元素规则
        :return:
        """
        return self.wait.until(ec.presence_of_element_located((By.XPATH, pattern)))

    def get_elements_by_xpath(self, pattern):
        """
        通过xpath获取元素
        :param pattern: 元素规则
        :return:
        """
        return self.wait.until(ec.presence_of_all_elements_located((By.XPATH, pattern)))

    def click_element_by_js(self, element):
        """
        通过js方式点击元素
        :param element:
        :return:
        """
        self.chrome.execute_script('arguments[0].click();', element)

    def login(self):
        """
        登录账号
        :return:
        """
        # 打开登录界面
        self.chrome.get(self.config['url'])
        # 输入账号
        self.wait.until(ec.presence_of_element_located((By.XPATH, '//*[@id="loginForm"]/div[1]/input'))).send_keys(
            self.config['username'])
        # 输入密码
        self.wait.until(ec.presence_of_element_located((By.XPATH, '//*[@id="loginForm"]/div[2]/input'))).send_keys(
            self.config['password'])
        # 点击登录
        self.wait.until(ec.presence_of_element_located((By.XPATH, '//*[@id="loginBtn"]'))).click()
        # 登录成功
        self.wait.until(ec.presence_of_element_located((By.XPATH, '/html/body/div[1]/a/img')))


class Page(Base):
    def __init__(self, config):
        super().__init__(config)
        self.my_data = []

    def get_excel_data(self):
        """
        获取已读取的excel数据
        :return:
        """
        excel = xlrd.open_workbook(self.config['excel_save_path'])
        sheet = excel.sheet_by_index(0)
        self.my_data = sheet.col_values(0)

    def find_result(self):
        """
        查找数据
        :return:
        """
        # 点击 我是分销商
        self.get_element_by_xpath('/html/body/div[3]/ul/li[3]').click()
        # 点击 可卖商品
        self.get_element_by_xpath('//*[@id="menu-3-2"]/ul/li[1]/a').click()
        # 切换iframe
        iframe = self.get_element_by_xpath('//*[@id="my_frame"]')
        self.chrome.switch_to.frame(iframe)
        # 点击 美国
        self.get_element_by_xpath('//*[@id="select-country-code"]/ul/li[3]/a').click()
        # 输入 最小库存
        self.get_element_by_xpath('//*[@id="minQnt"]').send_keys(self.config['min_qnt'])
        # 输入 最小供货价
        self.get_element_by_xpath('//*[@id="minSupPrice"]').send_keys(self.config['min_super_price'])
        # 点击 查询
        self.get_element_by_xpath('//*[@id="search-button"]').click()

    def get_all_page(self):
        """
        获取每页商品的页码
        :return:
        """
        while True:
            # 获取页面数据
            self.get_page_url()
            # 点击下一页
            next_btn = self.chrome.find_element_by_xpath('//div[@class="tcdPageCode"]/a[@class="nextPage"]')
            if not next_btn:
                break
            next_btn.click()
            time.sleep(5)

    def get_page_url(self):
        """
        获取当前页面所有的商品url
        :return:
        """
        goods_id_list, user_id_list = [], []
        retry_num = 10
        # 获取商品id和卖家id
        while retry_num:
            goods_id_list = self.get_elements_by_xpath('//*[@id="j-tbody"]/tr//td[3]')
            user_id_list = self.get_elements_by_xpath('//*[@id="j-tbody"]/tr//td[7]')
            if len(goods_id_list) == 100:
                break
            time.sleep(3)
            retry_num -= 1
        # 遍历商品 并 拼接url
        my_list = []
        for index in range(len(goods_id_list)):
            try:
                base_url = 'https://www.wanyilian.com/erp/winit_pro_view_new.php?pid={}&uname={}'
                pid, uname = goods_id_list[index].text[-5:], user_id_list[index].text
                url = base_url.format(pid, uname)
                if pid and uname:
                    # 若未爬取，则添加到数组中
                    if url not in self.my_data:
                        my_list.append(url)
            except:
                pass
        my_lock.acquire()
        urls.extend(my_list)
        my_lock.release()

    def run(self):
        """
        启动
        :return:
        """
        # 获取excel数据
        self.get_excel_data()
        # 登录
        self.login()
        # 查询数据
        self.find_result()
        # 爬取数据
        self.get_all_page()


class Details(Base):
    def __init__(self, config):
        super().__init__(config)
        # 保存数据的excel对象
        self.save_excel = openpyxl.load_workbook(self.config['excel_save_path'])
        self.save_sheet = self.save_excel[self.save_excel.sheetnames[0]]
        self.save_sheet['H1'] = '微信号：'
        self.save_sheet['H2'] = '19137599372'
        self.my_length = 0

    def get_all_data(self):
        """
        遍历每一个url获取详细数据
        :return:
        """
        while True:
            try:
                # 页码列表为0，页码线程已启动且页码线程已结束
                if len(urls) == 0:
                    if page_thread and not page_thread.is_alive():
                        break
                    else:
                        time.sleep(1)
                        continue
                my_lock.acquire()
                url = urls.pop()
                my_lock.release()
                self.get_details(url)
            except:
                pass

    def get_details(self, url):
        """
        获取商品详情
        :return:
        """
        while True:
            self.chrome.get(url)
            if self.chrome.current_url == url:
                break
            time.sleep(1)
        try:
            rows = self.get_elements_by_xpath('//*[@id="skuForm"]/table/tbody/tr')
            info = [
                self.get_element_by_xpath('//*[@id="form"]/div[4]/div/input').get_attribute('value'),
                url,
                self.get_element_by_xpath('//*[@id="skuForm"]/table/tbody/tr[1]/td[3]').text.replace('  正常商品', ''),
                self.get_element_by_xpath('//*[@id="skuForm"]/table/tbody/tr[1]/td[5]').text,
            ]
            if len(rows) == 1:
                info.append(self.get_element_by_xpath('//*[@id="skuForm"]/table/tbody/tr/td[last()-1]/div/p[3]').text)
            else:
                info.append(
                    self.get_element_by_xpath('//*[@id="skuForm"]/table/tbody/tr[last()]/td[last()-1]/div/p[3]').text)
            if info:
                info[-1] = info[-1].replace('尾程费用(请选择分区)：', '').replace('(USD)', '').replace('(GBP)', '')
                # 写入行
                self.save_sheet.append(info)
                # 保存
                self.save_excel.save(self.config['excel_save_path'])
                self.my_length += 1
                print('已爬取{}条，数据:{}'.format(self.my_length, info))
        except Exception as e:
            print(e)

    def run(self):
        """
        启动
        :return:
        """
        # 登录
        self.login()
        # 爬取数据
        self.get_all_data()


def run():
    config = {
        'username': '',
        'password': '',
        'url': 'http://www.wanyilian.com/erp/',
        'excel_save_path': './info.xlsx',
        'min_qnt': 5,
        'min_super_price': 100,
    }
    # 创建爬虫对象
    page = Page(config)
    details = Details(config)
    global page_thread, details_thread
    # 创建线程
    page_thread = Thread(target=page.run)
    details_thread = Thread(target=details.run)
    # 启动线程
    page_thread.start()
    details_thread.start()
    # 阻塞
    page_thread.join()
    details_thread.join()


if __name__ == '__main__':
    while True:
        run()
        time.sleep(3)
    # pyinstaller -F wanyilian.py
