#!/usr/bin/envpython3
# -*- coding: utf-8 -*-
"""规定编码"""
# pylint: disable=C0103
# pylint: disable=unused-variable, line-too-long, consider-using-enumerate
import os
import sys
import traceback
import time
from datetime import datetime
import pandas as pd
import win32api
import win32con
from selenium import webdriver
from selenium.common.exceptions import NoAlertPresentException, NoSuchElementException, ElementNotInteractableException
from selenium.webdriver.common.by import By
import requests
from PIL import Image
import threading
import auto_chromedrive as ac


def cut_screen(driver):
    """截图"""
    open_windows = driver.window_handles  # 获取打开的多个窗口句柄
    driver.switch_to.window(open_windows[-1])  # 切换到当前最新打开的窗口
    screen_name = time.strftime("%Y-%m-%d-%H.%M.%S") + '.png'
    driver.save_screenshot(screen_name)


def img_shot(driver):
    """保存考试二维码"""
    img_url = driver.find_element(By.ID, 'img_divCode').get_attribute("src")
    img_res = requests.get(img_url)
    img_name = driver.find_element(By.ID, 'lblExamName').text + '.png'
    img_file = open(img_name, 'wb')
    img_file.write(img_res.content)
    img_file.close
    img = Image.open(img_name)
    img.show()


def img_threading(threading_name, arg_name):
    """当打开图片时程序会暂停执行,利用子线程可以解决这个问题"""
    img_thread = threading.Thread(target=threading_name, args=arg_name)
    # img_thread.setDaemon(True)
    img_thread.start()
    time.sleep(2)  # 暂停等待子线程执行完毕


def sign_in(driver):
    """签到"""
    open_windows = driver.window_handles  # 获取打开的多个窗口句柄
    driver.switch_to.window(open_windows[-1])  # 切换到当前最新打开的窗口
    driver.get("http://u.sdhsg.com/sty/index.htm")
    open_windows = driver.window_handles  # 获取打开的多个窗口句柄
    driver.switch_to.window(open_windows[-1])  # 切换到当前最新打开的窗口
    try:
        driver.find_element(By.ID, 'a_sign').click()
    except NoSuchElementException:
        cut_screen(driver)
        print('签到失败,已截图')


def LearnCourse(driver):
    """从课程资源开始学习"""
    open_windows = driver.window_handles  # 获取打开的多个窗口句柄
    driver.switch_to.window(open_windows[-1])  # 切换到当前最新打开的窗口
    driver.refresh()  # 刷新页面，防止循环学习时‘已完成’标签不刷新
    time.sleep(1)  # 防止下面一句代码找不到elements
    while True:
        learn_courses = driver.find_elements(
            By.CLASS_NAME, 'el-kng-bottom-detail')
        for course in enumerate(learn_courses):
            course_object = course[1]
            course_name = course_object.text.split('\n')[0]
            course_num = course[0]
            Xpath_name = '//*[@id="SearchCatalog_Template_divTemplateHtmlContents"]/div/div[2]/div[3]/ul[2]/li[' + str(
                course_num + 1) + ']/div[1]'
            learn_state = course_object.find_element(By.XPATH, Xpath_name).text
            learn_verification = False
            if learn_state != '已完成':
                Xpath_parent = Xpath_name + '/..'  # 父元素的Xpath
                click_object = course_object.find_element(
                    By.XPATH, Xpath_parent)
                click_object.click()  # 点开课程包
                window_handles = driver.window_handles
                driver.switch_to.window(window_handles[-1])
                try:  # 点击开始学习
                    driver.find_element_by_css_selector(
                        "input#btnStartStudy").click()
                    try:  # 点开学习是考试的情况，关闭页面,继续往下遍历未学习课程
                        driver.find_element(By.ID, 'lblExamName')
                        print('课程 %s 考试未完成,遍历下一课程...' % course_name)
                        img_threading(img_shot, {driver: driver})
                        driver.close()
                        window_handles = driver.window_handles
                        driver.switch_to.window(window_handles[-1])
                        continue
                    except NoSuchElementException:
                        pass
                except NoSuchElementException:  # 如没有二级页面则不用再次点击
                    pass
                learn_verification = True
                print('开始学习: %s' % course_name)
                return None  # 跳出遍历
        if learn_verification is False:
            driver.find_element(By.LINK_TEXT, '下一页').click()
        if learn_verification is True:
            break


def auto_login(web_address, driver):
    """登录流程"""
    # 环境配置
    chromedriver = r'C:\Program Files (x86)\Google\Chrome\Application'
    os.environ["webdriver.ie.driver"] = chromedriver
    account_bool = os.path.exists('accountfile.csv')
    if account_bool is False:
        account_name = input('请输入账号:')
        account_password = input('请输入密码:')
        user_name = input('请输入用户名(自定义):')
        learn_time_input = int(input('请输入学习时长(单位:s):'))
        accountdata_generate = [
            [account_name, account_password, user_name, learn_time_input]]
        accountframe_generate = pd.DataFrame(
            accountdata_generate, columns=['账号', '密码', '用户', '学习时长'])
        accountframe_generate.to_csv('accountfile.csv', index=0)
    filepath = './accountfile.csv'  # 存储账号密码文件
    account_data = pd.read_csv(filepath)  # 读取账号密码
    accountname = account_data.iloc[0][0]
    password = account_data.iloc[0][1]
    learn_time = account_data.iloc[0][3]
    driver.get(web_address)  # 填单网站，未登录会显示登录
    driver.maximize_window()  # 最大化谷歌浏览器
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
    return learn_time


def learn_continue(driver, start_time, learn_time):
    """设定学习时长检查继续学习弹窗并点击"""
    open_windows = driver.window_handles  # 获取打开的多个窗口句柄
    driver.switch_to.window(open_windows[-1])  # 切换到当前最新打开的窗口
    time_now = datetime.now()
    time_learn_object = time_now - start_time
    time_learn_seconds = time_learn_object.seconds
    while time_learn_seconds < learn_time:
        time.sleep(100)  # 每隔100s处理一次弹窗
        try:  # 判断是否出现继续学习弹窗
            learn_button1 = driver.find_element(By.ID, "reStartStudy")
            learn_button1.click()
        except NoSuchElementException:
            pass
        try:
            learn_button2 = driver.find_element(
                By.XPATH, '//*[@id="dvHeartTip"]/input[2]')
            learn_button2.click()
        except NoSuchElementException:
            pass
        try:  # 判断是否跳转到考试页
            test_page = driver.find_element(By.ID, 'lblExamName')
            test_name = test_page.text
            print('课程 %s 需要进行考试' % test_name)
            img_threading(img_shot, {driver: driver})
            break
        except NoSuchElementException:
            pass
        try:  # 判断视频是否播放结束
            finish_element = driver.find_element(By.ID, "spanFinishContent")
            finish_text = '恭喜您已完成本视频的学习'
            if finish_text in finish_element.text:
                # time_now = datetime.now()
                # return time_now
                break
        except NoSuchElementException:
            pass
        try:  # 判断播放是否中断，出现加载对象
            refresh_element = driver.find_element(
                By.CSS_SELECTOR, '.jw-icon.jw-icon-display.jw-button-color.jw-reset')
            refresh_element.click()
        except ElementNotInteractableException:
            pass
        except NoSuchElementException:
            pass
        time_now = datetime.now()
        time_learn_object = time_now - start_time
        time_learn_seconds = time_learn_object.seconds
    time_now = datetime.now()
    time_learn_object = time_now - start_time
    time_learn_seconds = time_learn_object.seconds
    driver.close()
    return time_learn_seconds


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
    option.add_experimental_option(
        'excludeSwitches', ['enable-logging'])  # 处理一个错误提示信息
    driver_name = webdriver.Chrome(chrome_options=option)  # 选择Chrome浏览器
    # 登录课程资源网站
    source_website = 'http://u.sdhsg.com/kng/knowledgecatalogsearch.htm?sf=UploadDate&s=ac&st=null&mode='
    learn_time = auto_login(source_website, driver_name)
    start_time = datetime.now()
    print('开始学习,当前时间: %s' % start_time)
    print('学习时长: %ds' % learn_time)
    print('学习中...')
    learning_time = 0
    while learning_time < learn_time:
        LearnCourse(driver_name)
        learning_time = learn_continue(driver_name, start_time, learn_time)
    sign_in(driver_name)
    cut_screen(driver_name)
    driver_name.quit()
    print('学习完成,当前时间: %s' % datetime.now())
    currenttime = time.strftime('%m.%d %H:%M:%S', time.localtime())
    win32api.MessageBox(0, currenttime + "学习签到完成", "提醒", win32con.MB_OK)
