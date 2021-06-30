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

# 구글 크롤링 함수
def g_search(keyword, cnt):
    # 옵션 생성
    options = webdriver.ChromeOptions()
    # 창 숨기는 옵션 추가
    options.add_argument("headless")
    driver = webdriver.Chrome('C:/Users/LDCC/chromedriver.exe', options = options)
    # driver = webdriver.Chrome('C:/Users/LDCC/chromedriver.exe')
    driver.implicitly_wait(5)

    site = "https://www.google.com/search?q="+keyword+"&tbm=nws&sxsrf=ALeKk008u5T6jS5l0jLsN25hbj3J7Z0dvQ:1624256599096&source=lnt&tbs=qdr:d&sa=X&ved=2ahUKEwjRqPOsi6jxAhV4yosBHbRfAUUQpwV6BAgHECQ&biw=1920&bih=937&dpr=1"
    driver.get(site)
    req = driver.page_source
    soup = BeautifulSoup(req, 'html.parser')

    try:
        # wb = Workbook()
        # news = wb.active
        # news.title = "news"
        # news.append(["제목", "내용", "링크", "출처", "썸네일"])

        # 리스트 형태로 append (엑셀 파일 만들기 위해)
        news = []
        for i in range(cnt):
            source = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "#rso > div:nth-child("+str(i+1)+") > g-card > div > div > div.dbsr > a > div > div.hI5pFf > div.XTjFC.WF4CUc"))
            )

            source_time = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "#rso > div:nth-child("+str(i+1)+") > g-card > div > div > div.dbsr > a > div > div.hI5pFf > div.yJHHTd > div.wxp1Sb > span > span > span"))
            )

            title = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "#rso > div:nth-child("+str(i+1)+") > g-card > div > div > div.dbsr > a > div > div.hI5pFf > div.JheGif.nDgy9d"))
            )

            url = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "#rso > div:nth-child("+str(i+1)+") > g-card > div > div > div.dbsr > a"))
            )

            content = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "#rso > div:nth-child("+str(i+1)+") > g-card > div > div > div.dbsr > a > div > div.hI5pFf > div.yJHHTd > div.Y3v8qd"))
            )

            news.append([title.text, content.text, url.get_attribute("href"), source.text + " " + source_time.text])

        # list to dataframe
        news_df = pd.DataFrame(news, columns = ["title", "content", "url", "source"])

        return news_df

    finally:
        driver.quit()

# 네이버 크롤링 함수
def n_search(keyword, cnt):
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
        for i in range(cnt):
            if "네이버뉴스" in sources[i].text:
                news.append([titles[i].text, contents[i].text, titles[i].get_attribute("href"), sources[i].text[:-5], imgs[i].get_attribute("src")])
                continue

            news.append([titles[i].text, contents[i].text, titles[i].get_attribute("href"), sources[i].text, imgs[i].get_attribute("src")])

        # list to dataframe
        news_df = pd.DataFrame(news, columns = ["title", "content", "url", "source", "thumbnail"])

        return news_df

    finally:
        driver.quit()
        # wb.save(filename='news.xlsx')


# 다음 크롤링 함수
def d_search(keyword, cnt):
    # 옵션 생성
    options = webdriver.ChromeOptions()
    # 창 숨기는 옵션 추가
    options.add_argument("headless")
    driver = webdriver.Chrome('C:/Users/LDCC/chromedriver.exe', options = options)
    # driver = webdriver.Chrome('C:/Users/LDCC/chromedriver.exe')
    driver.implicitly_wait(5)

    site = "https://search.daum.net/search?w=news&nil_search=btn&DA=STC&enc=utf8&cluster=y&cluster_page=1&q="+str(keyword)+"&period=d&sd=20210620172106&ed=20210621172106"
    driver.get(site)
    req = driver.page_source
    soup = BeautifulSoup(req, 'html.parser')

    try:
        # 리스트 형태로 append (엑셀 파일 만들기 위해)
        news = []
        for i in range(cnt):
            source = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="newsColl"]/div[1]/ul/li['+str(i+1)+']/div[2]/span[1]'))
            )

            title = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="newsColl"]/div[1]/ul/li['+str(i+1)+']/div[2]/a'))
            )

            content = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="newsColl"]/div[1]/ul/li['+str(i+1)+']/div[2]/p[2]'))
            )

            news.append([title.text, content.text, title.get_attribute("href"), source.text])

        # list to dataframe
        news_df = pd.DataFrame(news, columns = ["title", "content", "url", "source"])

        return news_df

    finally:
        driver.quit()


# 실행 코드
if __name__ == "__main__":
    keyword = "롯데"
    cnt = 3

    g_news_df = g_search(keyword = keyword, cnt = cnt)  # "title", "content", "url", "source"
    n_news_df = n_search(keyword = keyword, cnt = cnt)      # "title", "content", "url", "source"
    d_news_df = d_search(keyword = keyword, cnt = cnt)      # "title", "content", "url", "source"

    to = ["hansol.ko@lotte.net"]
    # to = ["hansol.ko@lotte.net", "hyereen.kong@lotte.net"]
    subject = "[정보] 롯데 관련 NEWS " + "(20" + datetime.today().strftime("%y.%m.%d") + ")"
    g_content = "<h1>[구글 뉴스]</h1>"
    n_content = "<h1>[네이버 뉴스]</h1>"
    d_content = "<h1>[다음 뉴스]</h1>"

    for i in range(cnt):
        g_cont = """
        <h2>{}</h2>
        <blockquote><small>{}</small></blockquote>
        <p>{}</p>
        <p><em>출처</em> : {}</p>
        <hr>
        """.format(g_news_df["title"][i], g_news_df["source"][i], g_news_df["content"][i], g_news_df["url"][i])
        g_content += g_cont

    for i in range(cnt):
        n_cont = """
        <h2>{}</h2>
        <blockquote><small>{}</small></blockquote>
        <p>{}</p>
        <p><em>출처</em> : {}</p>
        <hr>
        """.format(n_news_df["title"][i], n_news_df["source"][i], n_news_df["content"][i], n_news_df["url"][i])
        n_content += n_cont

    for i in range(cnt):
        d_cont = """
        <h2>{}</h2>
        <blockquote><small>{}</small></blockquote>
        <p>{}</p>
        <p><em>출처</em> : {}</p>
        <hr>
        """.format(d_news_df["title"][i], d_news_df["source"][i], d_news_df["content"][i], d_news_df["url"][i])
        d_content += d_cont

    # 구글, 네이버, 다음 합치기
    final_content = g_content + n_content + d_content

    # 구글, 네이버, 다음 기사 합쳐서 메일 보내기 (1일 이내 최신 기사)
    for t in to:
        send_mail(t, subject, final_content)