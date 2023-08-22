import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from webdriver_manager.chrome import ChromeDriverManager
import time
from selenium.webdriver.chrome.options import Options
##參考資訊:https://www.learncodewithmike.com/2020/02/python-beautifulsoup-web-scraper.html
link =pd.read_excel(r"D:\代碼\網站清單.xlsx")
urlname = list(link['網站說明'])
#print(urlname)
url = link['管理網站 URL']
finalurl=[url[i] for i in range(len(url)) if urlname[i]=='司法院判決書系統'][0]

#利用Selenium開啟chrome 等待網頁載入完成
options = Options()
options.add_argument("--disable-notifications")  #不啟用通知

driver = webdriver.Chrome('Applications/chromedriver')
driver.get(finalurl)
time.sleep(10)

soup = BeautifulSoup(driver.page_source,'html.parser')###解析網頁
print(soup.title.getText())# 輸出網頁的 title
print(soup.prettify())  #輸出排版後的HTML內容
##選擇審判法院以及判決字號
class judge:
    def __init__(self, *jud_years):
        self.jud_years = jud_years
    def _init_(self,*jud_cases):
        self.jud_cases = jud_cases
    def _init_(self,*jud_nos):
        self.jud_nos = jud_nos
    def daily(self, jud_court):
        browser = webdriver.Chrome(ChromeDriverManager().install())
        browser.get(finalurl)
    
        select_court = Select(browser.find_element("jud_court"))
        select_court.select_by_value(jud_court)  #選擇審判的法院

        year = browser.find_element_by_name("jud_year")  # 定位判決年度
        case = browser.find_element_by_name("jud_case")  # 定位判決類型
        number = browser.find_element_by_name("jud_no")  # 定位判決類型

        result = []
        for jud_year in self.jud_years:
            year.clear()  # 清空判決年度
            year.send_keys(jud_year)
            year.submit()
            time.sleep(2)
        for jud_case in self.jud_cases:
            case.clear()  # 清空判決類型
            case.send_keys(jud_case)
            case.submit()
            time.sleep(2)
        for jud_no in self.jud_nos:
            number.clear()  # 清空判決暗號
            number.send_keys(jud_no)
            number.submit()
            time.sleep(2)
     

            soup = BeautifulSoup(browser.page_source, "lxml")
            table = soup.find("table", {"id": "jud"})
            elements = table.find_all("td", {"class": "hlTitle_scroll"})
            
            data = (jud_case,jud_case,jud_no,) + tuple(element.getText() for element in elements)
            
        result.append(data)


judge = judge("112", "訴","2511") 
judge.daily("台灣台中地方法院")