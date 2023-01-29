import math
import time
import pandas as pd
import requests
from bs4 import BeautifulSoup
import os

def GetDataFrameStocksPrice(codeList, pages):

    df = pd.DataFrame()
    for code in codeList:
        # 총 페이지 수 구하기
        urlTemp = "https://finance.naver.com/item/sise_day.naver?code={}&page=1".format(code)
        html = BeautifulSoup(requests.get(urlTemp, headers={'User-agent': 'Mozilla/5.0'}).text, 'lxml')
        pgrr = html.find('td', class_='pgRR')
        s = str(pgrr.a['href']).split('=')
        lastPage = s[-1]

        for page in range(1, pages):
            url = "https://finance.naver.com/item/sise_day.naver?code={}&page={}".format(code, page)
            res = requests.get(url, headers={'User-agent': 'Mozilla/5.0'})
            df = pd.concat([df, pd.read_html(res.text, header=0)[0]], axis=0)

        # df.dropna()를 이용해 결측값 있는 행 제거
        df = df.dropna()
    return df

def GetDataFrameStockPrice(code, pages):

    df = pd.DataFrame()
    # 총 페이지 수 구하기
    urlTemp = "https://finance.naver.com/item/sise_day.naver?code={}&page=1".format(code)
    html = BeautifulSoup(requests.get(urlTemp, headers={'User-agent': 'Mozilla/5.0'}).text, 'lxml')
    pgrr = html.find('td', class_='pgRR')
    s = str(pgrr.a['href']).split('=')
    lastPage = s[-1]

    for page in range(1, pages):
        url = "https://finance.naver.com/item/sise_day.naver?code={}&page={}".format(code, page)
        res = requests.get(url, headers={'User-agent': 'Mozilla/5.0'})
        df = pd.concat([df, pd.read_html(res.text, header=0)[0]], axis=0)
        df['저가-시가'] = df['저가']-df['시가']
        df['고가-시가'] = df['고가']-df['시가']

    # df.dropna()를 이용해 결측값 있는 행 제거
    df = df.dropna()
    # df = df[df["날짜"] != ""]
    return df

    # 날짜, 종가, 전일비, 시가, 고가, 저가, 거래량

def SelectDayBeforeStockPrice(code, beforeDate, gubun):

    if beforeDate == None: # ~일 전 데이터 없을 시, 최신일로 지정
        beforeDate = 0
    if gubun == "날짜":
        idxGubun = 0
    elif gubun == "종가":
        idxGubun = 1
    elif gubun == "전일비":
        idxGubun = 2
    elif gubun == "시가":
        idxGubun = 3
    elif gubun == "고가":
        idxGubun = 4
    elif gubun == "저가":
        idxGubun = 5
    elif gubun == "거래량":
        idxGubun = 6
    elif gubun == "저가-시가":
        idxGubun = 7
    elif gubun == "고가-시가":
        idxGubun = 8

    df = GetDataFrameStockPrice(code, 2).iat[beforeDate, idxGubun]

    return df


codelist = ['900110', '900270', '900300', '000250']

#파일 새로 작성
# for code in codelist:
#     GetDataFrameStockPrice(code, 10).to_excel('portfolio.xlsx', sheet_name=code, index = False, header=True)

# 파일 존재시
for code in codelist:
    with pd.ExcelWriter("portfolio.xlsx", mode='a', engine = 'openpyxl') as writer:
        GetDataFrameStockPrice(code, 2).to_excel(writer, sheet_name=code)

# for code in codelist:
#     df = GetDataFrameStockPrice(code, 10)
#     print(round((int(df.iat[0, 4]) - int(df.iat[1, 5])) / int(df.iat[0, 1]) * 100, 2))
