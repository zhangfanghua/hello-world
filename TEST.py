# -*- coding: UTF-8 -*-
from selenium import webdriver#selenium 為爬蟲的核心 用來啟動瀏覽器
from selenium.webdriver.common.keys import Keys #模擬鍵盤輸入的插件
from bs4 import BeautifulSoup#使用BeautifulSoup來解析網頁
from time import sleep #sleep為讓程式停止的意思 通常都是用在點擊後網頁會lag 就讓程式先停止稍後在做
import time, datetime #時間的插件
import numpy as np#資料的事前處理，多重維度的陣列
import pandas as pd#資料的事前處理，如資料補值，空值去除或取代等
import csv #為處理CSV的插件
import pymysql
from sqlalchemy import create_engine
import os #處理像是 檔案重新命名、移動、刪除......

chromedriver = r"\\vm02mfg\e$\VM02M0\資訊共享\Chromedriver\V91\chromedriver.exe"
options = webdriver.ChromeOptions()
### 下載路徑----------
prefs = {'profile.default_content_settings.popups': 0,  "download.default_directory": 'U:\\資料\\python\\tes\\'}
options.add_experimental_option('prefs', prefs)
options.add_argument("--disable-notifications")
options.add_experimental_option('useAutomationExtension', False)
options.add_experimental_option("excludeSwitches", ["enable-automation"])
browser = webdriver.Chrome(options=options,executable_path=chromedriver)#webdriver.Chrome為瀏覽器的驅動
#並且為他命名為browser
#以上不需要做更改

#上面開啟的瀏覽器卻沒有給它網址，現在我要給他一個網址讓它執行
browser.get("http://vm02mfg/default.aspx")#在" "中間填入要開啟的網址!

### 安全驗證-------------------------------------------------------------
try:
    browser.find_element_by_xpath('//*[@id="details-button"]').click()
    browser.find_element_by_xpath('//*[@id="proceed-link"]').click()
except:
    pass
#開啟問卷的網址，但它需要填寫帳號密碼，就需要用到模擬輸入
#---------------------------------------------------------------------------

#現在開始要模擬狀態，從登入帳號~點擊下拉式選單到送出都是要模擬一步一步來執行的

#click為點擊
#send.keys為模擬輸入


# browser.find_element_by_xpath('//*[@id="edit-name"]').send_keys("2107127")#點擊帳號的Xpath並且填入帳號!
# browser.find_element_by_xpath('//*[@id="edit-pass"]').send_keys("Kelly111")#點擊密碼的Xpath並且填入密碼!
# browser.find_element_by_xpath('//*[@id="edit-submit"]').click()#點擊送出的Xpath!
# #點擊送出的按鈕

# sleep(3)#因為網頁開比較慢，讓程式停三秒再執行
# browser.find_element_by_xpath('//*[@id="main-content"]/div/div[1]/ul[1]/li[2]/a').click()#!
# time.sleep(3)
# browser.find_element_by_xpath('//*[@id="main-content"]/div/div[1]/ul[2]/li[4]/a').click()#!
# time.sleep(3)
# browser.find_element_by_xpath('//*[@id="edit-submit"]').click()#!
# time.sleep(3)
# browser.close()
# browser.quit()
# # # ### 讀取 Excel dateframe 整理 ------------------------------------------#!
# df = pd.read_excel(r'U:\資料\python\tes\m02_temporary_leave_application_questionnaire_m02.xlsx',skiprows = 2,usecols= [2,6,8,10,11,12,13,14] ,encoding='utf-8') #,11,12,13,14,15,16,17,18,19,20,21,22]
# text = df[df['ID 工號'].astype(str).str.contains('[current-user:profile-sector:field_profile_emp_name]')]
# for index,row in df.iterrows():
#     # print(row['工號'].str.contains('[current-user:profile-sector:field_profile_emp_name]'))
#     # if type(row['工號']) == s
#     # print(type(row['工號']),row['工號'])
#     if   type(row['ID 工號']) ==  str :
#         # print(index)
#         # print(row['工號'],row['使用者名稱'])
#         # df['工號'].replace(row['工號'],row['使用者名稱'],inplace = True)
#         df['ID 工號'].iloc[index] = row['使用者名稱']
#     else :
#         pass
# del df['使用者名稱']

# ## 工號補0

# df['ID 工號'] = df['ID 工號'].apply(lambda x : '{:0>7d}'.format(x))

# ### 過濾重複填寫，篩選最新填寫資料
# df['ranks'] = df.groupby(['ID 工號'])['時間'].rank(method = 'min')
# df =df[df['ranks'] == 1]

# conn = pymysql.connect(host="10.30.163.205",user="admin",passwd="Vm02+1234",database='tableau')
# cursor = conn.cursor()
# sql = 'SELECT distinct workid  ,workname  , department    FROM tableau.h_104'

# df1 = pd.read_sql(sql,con = conn)
# df1.columns =  ['ID 工號','Name 姓名','Department 部門']
# df2 = pd.merge(df,df1)
# del df2['ranks']
# col = ['ID 工號','Name 姓名','Department 部門','Leave Type 請假假別','Leave Reason 請假理由1','Reason of feeling unwell 身體不適的原因','Leave Reason 請假理由2']
# df2 = df2[col]
# # print(df2)
# df2.to_excel(r'測試請假問卷.xlsx',index = False)

