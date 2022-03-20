from pickle import NONE
from timeit import repeat
from matplotlib.pyplot import hist
import pandas as pd
import time
import os
import sys
from selenium import webdriver
#先安装pywin32，才能导入下面两个包
import win32api
import win32con
#导入处理alert所需要的包
from selenium.common.exceptions import NoAlertPresentException, NoSuchElementException
import traceback
import win32api,win32con
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.keys import Keys
import auto_chromedrive as ac
import re
from selenium.webdriver.common.by import By
from autologin import course_data, auto_login, sel_address, history_file

# 获取当前程序目录的绝对路径和名称
work_path,work_name = os.path.split(os.path.abspath(sys.argv[0]))
# 更改当前工作目录
os.chdir(work_path)
# 更新浏览器驱动
web_address = 'http://u.sdhsg.com/apps/sty/mystudyhistoyimageinfonew.htm'
ac.checkChromeDriverUpdate()
driver_name = webdriver.Chrome()  # 选择Chrome浏览器
auto_login(web_address, driver_name)
history_file(driver_name)
