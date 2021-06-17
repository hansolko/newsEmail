# -*- coding: utf-8 -*-
import win32com.client
import time

from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from urllib.request import urlretrieve   # 이미지 다운로드
from openpyxl import Workbook   # 엑셀로 저장
import schedule
from datetime import datetime
import pandas as pd

# 메일 보내기 함수
def send_mail(to, subject, content, atch=[]):
    # Outlook Object Model 불러오기
    new_Mail = win32com.client.Dispatch("Outlook.Application").CreateItem(0)

    # 메일 수신자
    new_Mail.To = to
    # 메일 참조
    # new_Mail.CC = "mail-add-for-cc@testadd.com"
    # 메일 제목
    new_Mail.Subject = subject
    # 메일 내용
    new_Mail.HTMLBody = content

    # 첨부파일 추가
    if atch:
        for file in atch:
            new_Mail.Attachments.Add(file)

    # 메일 발송
    new_Mail.Send()

# 크롤링 함수
def search(keyword):
    # 옵션 생성
    options = webdriver.ChromeOptions()
    # 창 숨기는 옵션 추가
    options.add_argument("headless")
    driver = webdriver.Chrome('C:/Users/LDCC/chromedriver.exe', options = options)
    # driver = webdriver.Chrome('C:/Users/LDCC/chromedriver.exe')
    driver.implicitly_wait(5)

    site = "https://search.naver.com/search.naver?where=news&sm=tab_jum&query="+keyword+"&pd=4"
    driver.get(site)
    req = driver.page_source
    soup = BeautifulSoup(req, 'html.parser')

    try:
        # wb = Workbook()
        # news = wb.active
        # news.title = "news"
        # news.append(["제목", "내용", "링크", "출처", "썸네일"])

        elements = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "#main_pack > section.sc_new.sp_nnews._prs_nws > div > div.group_news > ul"))
        )
        sources = elements.find_elements_by_class_name('info_group')
        titles = elements.find_elements_by_class_name('news_tit')
        contents = elements.find_elements_by_class_name('news_dsc')
        imgs = elements.find_elements_by_css_selector('img.thumb.api_get')

        # 리스트 형태로 append (엑셀 파일 만들기 위해)
        news = []
        for i in range(len(sources)):
            news.append([titles[i].text, contents[i].text, titles[i].get_attribute("href"), sources[i].text, imgs[i].get_attribute("src")])

        # list to dataframe
        news_df = pd.DataFrame(news, columns = ["title", "content", "url", "source", "thumbnail"])

        return news_df

    finally:
        driver.quit()
        # wb.save(filename='news.xlsx')

# img_folder_path = 'C:/Users/LDCC/data_scientist/imgs'  # 이미지 저장 폴더
#
# if not os.path.isdir(img_folder_path):  # 없으면 새로 생성
#     os.mkdir(img_folder_path)
#
# for index, link in enumerate(img_url):           #리스트에 있는 원소만큼 반복, 인덱스는 index에, 원소들은 link를 통해 접근 가능
#     start = link.rfind('.')         #.을 시작으로
#     end = link.rfind('?')           #?를 끝으로
#     filetype = link[start:end]      #확장자명을 잘라서 filetype변수에 저장 (ex -> .jpg)
#     urlretrieve(link, 'C:/Users/LDCC/data_scientist/imgs/{}.jpg'.format(index))        #link에서 이미지 다운로드, './imgs/'에 파일명은 index와 확장자명으로

# 실행 코드
if __name__ == "__main__":
    news_df = search(keyword="롯데")      # "title", "content", "url", "source", "thumbnail"

    to = "hansol.ko@lotte.net"
    subject = "[정보] 롯데 관련 NEWS " + "(20" + datetime.today().strftime("%y.%m.%d") + ")"
    content = ""
    for i in range(len(news_df)):
        cont = """
        <h1>{}</h1>
        <blockquote><small>{}</small></blockquote>
        <p>{}</p>
        <p><em>출처</em> : {}</p>
        <hr>
        """.format(news_df["title"][i], news_df["source"][i], news_df["content"][i], news_df["url"][i])
        content += cont

    send_mail(to, subject, content)