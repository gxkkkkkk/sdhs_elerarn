#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""规定编码"""
# pylint: disable=C0103
# pylint: disable=unused-variable, line-too-long, consider-using-enumerate
import os
import sys
import re
import traceback
import math
import time
import pandas as pd
from sympy import false
import win32api
import win32con
from selenium import webdriver
from selenium.common.exceptions import NoAlertPresentException, NoSuchElementException, StaleElementReferenceException, TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import auto_chromedrive as ac


def organize_list(history_list):
    """函数history_file的附属函数,方便处理数据"""
    for word in enumerate(history_list):
        re_name = re.match(r'学习 \d{1,}分钟', word[1])
        if word[0] > 0:
            test_num = word[0] - 1
            test_name = '课程包' in history_list[test_num]
            if (re_name is not None) and (test_name is False):
                history_list.insert(word[0], '课程包:' + history_list[test_num])
                organize_list(history_list)


def load_more(driver):
    """点击学习历史中的加载更多,方便爬取数据和点击学习"""
    button_load = driver.find_element(By.CSS_SELECTOR, 'div#addMore')
    load_text = button_load.text
    compare_text = '加载更多'
    if load_text == compare_text:
        try:
            button_load.click()
            load_more(driver)
        except StaleElementReferenceException:
            time.sleep(5)
            load_more(driver)


def sign_in(driver):
    """签到"""
    open_windows = driver.window_handles  # 获取打开的多个窗口句柄
    driver.switch_to.window(open_windows[-1])  # 切换到当前最新打开的窗口
    driver.get("http://u.sdhsg.com/sty/index.htm")
    open_windows = driver.window_handles  # 获取打开的多个窗口句柄
    driver.switch_to.window(open_windows[-1])  # 切换到当前最新打开的窗口
    driver.find_element(By.ID, 'a_sign').click()


def course_data(driver):
    """提取全站资源课程信息"""
    course_text_data = [[], [], [], []]
    while True:
        course_list = driver.find_elements(
            By.CLASS_NAME, 'el-kng-bottom-detail')
        print(len(course_list))
        while len(course_list) == 0:
            course_list = driver.find_elements(
                By.CLASS_NAME, 'el-kng-bottom-detail')
            # print(f'修正{len(course_list)}')
        for course in course_list:
            course_text = course.text.split('\n')
            for course_num in range(len(course_text_data)):
                if len(course_text) == 3:
                    course_text.insert(1, None)
                course_text_data[course_num].append(course_text[course_num])
        try:
            driver.find_element(By.LINK_TEXT, '下一页').click()
        except NoSuchElementException:  # 处理最后一页无法点击下一页的错误
            break
    course_df = pd.DataFrame()
    course_df.insert(loc=0, column='资源名称', value=course_text_data[0])
    course_df.insert(loc=1, column='发布人员', value=course_text_data[1])
    course_df.insert(loc=2, column='发布时间', value=course_text_data[2])
    course_df.insert(loc=3, column='学习人数', value=course_text_data[3])
    course_df.to_excel('课程资源数据.xlsx', index=False)


def history_file(driver):
    """生成学习学习历史数据文件"""
    history_bool = os.path.exists('学习历史数据.xlsx')
    open_windows = driver.window_handles  # 获取打开的多个窗口句柄
    driver.switch_to.window(open_windows[-1])  # 切换到当前最新打开的窗口
    driver.get('http://u.sdhsg.com/apps/sty/mystudyhistoyimageinfonew.htm')
    open_windows = driver.window_handles  # 获取打开的多个窗口句柄
    driver.switch_to.window(open_windows[-1])  # 切换到当前最新打开的窗口
    if history_bool is False:
        load_more(driver)
    learn_history = driver.find_element_by_class_name(
        'el-study-history-list')  # 获取学习历史数据
    history_list = learn_history.text.splitlines()  # 获取文本
    # 整理文本
    organize_list(history_list)
    for element_word in history_list:
        if re.match(r'\d{4}-\d{2}-\d{2}', element_word) is not None:
            history_list.remove(element_word)  # 移除日期
    for element_word in enumerate(history_list):
        num_bundle = element_word[0]
        next_num = num_bundle + 1
        text_bundle = element_word[1]
        if '课程包' in text_bundle:
            text_deal = text_bundle[4:]
            history_list[num_bundle] = text_deal
        if ('分钟' in text_bundle) and ('%' not in history_list[next_num]):
            history_list.insert(next_num, '100%')
    list_exchange = [[], [], [], []]
    for element_word in enumerate(history_list):
        word_num = element_word[0]
        word_content = element_word[1]
        if (word_num % 4) == 0:
            list_exchange[0].append(word_content)
        elif (word_num % 4) == 1:
            list_exchange[1].append(word_content)
        elif (word_num % 4) == 2:
            list_exchange[2].append(word_content)
        elif (word_num % 4) == 3:
            list_exchange[3].append(word_content)
    history_df = pd.DataFrame()
    history_df.insert(loc=0, column='资源名称', value=list_exchange[0])
    history_df.insert(loc=1, column='课程包', value=list_exchange[1])
    history_df.insert(loc=2, column='学习时长', value=list_exchange[2])
    history_df.insert(loc=3, column='学习进度', value=list_exchange[3])
    history_df = history_df.drop_duplicates('资源名称')  # 删除重复数据，只保留最新的记录
    history_df = history_df.reset_index(drop=True)
    if history_bool is True:
        history_before = pd.read_excel('学习历史数据.xlsx', header=0)
        history_later = pd.merge(history_df, history_before, how='outer')
        history_later.to_excel('学习历史数据.xlsx', index=False)
    else:
        history_df.to_excel('学习历史数据.xlsx', index=False)
        driver.quit()


def read_history(driver):
    """读取历史学习数据，提取未学完的资源名"""
    history_bool = os.path.exists('学习历史数据.xlsx')
    if history_bool is False:
        web_address = 'http://u.sdhsg.com/apps/sty/mystudyhistoyimageinfonew.htm'
        auto_login(web_address, driver)
        history_file(driver)
    history_data = pd.read_excel('学习历史数据.xlsx')
    rate_list = history_data['学习进度']
    learn_name = None
    for rate_number in enumerate(rate_list):
        rate_num = rate_number[0]
        rate = rate_list[rate_num]
        if float(rate.strip('%')) < 100:
            learn_name = history_data['资源名称'][rate_num]
            break
    return learn_name


def sel_address(learn_name):
    """选择登录网站"""
    if learn_name is None:
        web_address = 'http://u.sdhsg.com/kng/knowledgecatalogsearch.htm?sf=UploadDate&s=ac&st=null&mode='
    else:
        web_address = 'http://u.sdhsg.com/apps/sty/mystudyhistoyimageinfonew.htm'
    return web_address


def learn_for_history(driver, plan_name):
    """从学习历史开始学习"""
    # load_more(driver)
    learn_data = driver.find_element_by_class_name('el-study-history-list')
    learn_list = learn_data.find_elements_by_tag_name('ul')
    for li_number in enumerate(learn_list):
        number_li = li_number[0]
        repeat_num = 1
        li_data = learn_list[number_li].find_elements_by_tag_name('li')
        for li in li_data:
            source_name = li.text
            if plan_name in source_name:
                repeat_num = 0
                div_list = li.find_elements_by_tag_name('div')
                for div in div_list:
                    if (plan_name in div.text) and ('课程包' not in div.text):
                        div.click()
                        break
                break
        if repeat_num == 0:
            break


class LearnCourse():
    """从课程资源开始学习,对照学习历史数据,不点击学习进度100%的课程"""

    def __init__(self, driver_option):
        history_learn = pd.read_excel('学习历史数据.xlsx')
        finish_learn = history_learn[history_learn['学习进度'] == '100%']
        self.finish_name = finish_learn['课程包'].to_list()
        self.driver = driver_option

    def learn_click(self):
        """从课程资源开始学习"""
        time.sleep(1)  # 防止下面一句代码找不到elements
        learn_courses = self.driver.find_elements(
            By.CLASS_NAME, 'el-kng-bottom-detail')
        for course in enumerate(learn_courses):
            course_name = course[1].text.split('\n')[0]
            course_num = course[0]
            # self.finish_name.append(course_name)
            if course_name not in self.finish_name:
                course[1].find_element_by_css_selector(
                    ".font-size-14.hand.text-normal").click()  # 点击课程
                window_handles = self.driver.window_handles
                self.driver.switch_to.window(window_handles[-1])
                try:  # 点击开始学习
                    self.driver.find_element_by_css_selector(
                        "input#btnStartStudy").click()
                except NoSuchElementException:  # 如没有二级页面则退出
                    pass
                break
            elif course_num == 19:  # 遍历下一页
                self.driver.find_element(By.LINK_TEXT, '下一页').click()
                self.learn_click()


def auto_login(web_address, driver):
    """登录流程"""
    # 环境配置
    chromedriver = r'C:\Program Files (x86)\Google\Chrome\Application'
    os.environ["webdriver.ie.driver"] = chromedriver
    filepath = './accountfile.csv'  # 存储账号密码文件
    account_data = pd.read_csv(filepath)  # 读取账号密码
    accountname = account_data.iloc[0][0]
    password = account_data.iloc[0][1]
    driver.get(web_address)  # 填单网站，未登录会显示登录
    # driver.maximize_window() #最大化谷歌浏览器
    # 处理alert弹窗
    try:
        alert1 = driver.switch_to.alert  # switch_to.alert点击确认alert
    except NoAlertPresentException:
        print("no alert")
        traceback.print_exc()
    else:
        at_text1 = alert1.text
        print("at_text:" + at_text1)
    username = str(accountname)
    password = str(password)
    driver.implicitly_wait(1)  # 设置隐式等待
    driver.find_element_by_id('txtUserName2').click()  # 点击用户名输入框
    driver.find_element_by_id('txtUserName2').clear()  # 清空输入框
    driver.find_element_by_id('txtUserName2').send_keys(username)  # 自动敲入用户名
    driver.find_element_by_id('txtPassword2').click()  # 点击密码输入框
    driver.find_element_by_id('txtPassword2').clear()  # 清空输入框
    driver.find_element_by_id('txtPassword2').send_keys(password)  # 自动敲入密码
    # 采用xpath定位登陆按钮
    driver.find_element_by_xpath('//*[@id="btnLogin2"]').click()
    error_element = driver.find_element(By.ID, 'lblErrMsg1')
    error_text = error_element.text
    if error_text == '请阅读并同意下方隐私协议':
        driver.find_element(By.ID, 'chkloginpass').click()  # 点击勾选同意协议
        driver.find_element_by_xpath('//*[@id="btnLogin2"]').click()  # 点击登录


def learn_continue(driver, time_learn):
    """设定学习时长检查继续学习弹窗并点击"""
    time_one = 600  # 确切弹窗时间
    time_redundant = 700  # 设定冗余弹窗时间
    time_popup = time_one + time_redundant  # 最长等待时间
    times = math.ceil(time_learn * 60 / time_one) - 1  # 点击次数
    open_windows = driver.window_handles  # 获取打开的多个窗口句柄
    driver.switch_to.window(open_windows[-1])  # 切换到当前最新打开的窗口
    for i in range(times):  # 循环点击
        try:
            learn_warning = WebDriverWait(driver, time_popup, poll_frequency=60, ignored_exceptions=None).until(
                EC.presence_of_element_located((By.ID, "reStartStudy")))
        except TimeoutException:
            finish_element = driver.find_element(By.ID, "spanFinishContent")
            finish_text = '恭喜您已完成本视频的学习'
            if finish_text in finish_element.text:
                break
        else:
            driver.find_element(By.ID, "reStartStudy").click()
    driver.close()


if __name__ == '__main__':
    # 获取当前程序目录的绝对路径和名称
    work_path, work_name = os.path.split(os.path.abspath(sys.argv[0]))
    # 更改当前工作目录
    os.chdir(work_path)
    # 更新浏览器驱动
    ac.checkChromeDriverUpdate()
    option = webdriver.ChromeOptions()
    # option.add_argument('headless')  # 隐藏浏览器
    option.add_argument("--mute-audio")  # 静音
    driver_name = webdriver.Chrome(chrome_options=option)  # 选择Chrome浏览器
    history_bool = os.path.exists('学习历史数据.xlsx')  # 检查是否存在学习历史数据
    old_learn = read_history(driver_name)
    sel_web = sel_address(old_learn)
    if history_bool is False:
        option = webdriver.ChromeOptions()
        # option.add_argument('headless')  # 隐藏浏览器
        option.add_argument("--mute-audio")  # 静音
        driver_name = webdriver.Chrome(chrome_options=option)  # 选择Chrome浏览器
    auto_login(sel_web, driver_name)
    learning_time = 60  # 设定学习时长，单位：分钟
    if old_learn is not None:
        learn_for_history(driver_name, old_learn)
        learn_continue(driver_name, learning_time)
    else:
        learn_course = LearnCourse(driver_name)
        learn_course.learn_click()
        learn_continue(driver_name, learning_time)
    time.sleep(60)  # 等候60s
    history_file(driver_name)
    sign_in(driver_name)
    driver_name.quit()
    currenttime = time.strftime('%m.%d %H:%M:%S', time.localtime())
    win32api.MessageBox(0, currenttime + "签到完成", "提醒", win32con.MB_OK)
