# stx, stp 환율크롤링 및 db 연동
### 크롤링 사이트 : https://www.bnm.gov.my/exchange-rates/ 
### 　　　　　　　http://www.safe.gov.cn/safe/rmbhlzjj/index.html 




## 목차 
     1. 환율 크롤링
     2. ORACLE DB 접속 및 INSERT
     3. Notification
     4. Scheduling

## 1.환율 크롤링

```py
# stx 환율데이터
import re
import requests
import time
from bs4 import BeautifulSoup
import lxml
from outlook_mail import outlook_mail_web
from decimal import Decimal

def STX_currency():

    url = 'http://www.safe.gov.cn/AppStructured/hlw/RMBQuery.do'
    res = requests.get(url, timeout=20)
    if res:
        soup = BeautifulSoup(res.content, 'lxml')

        test = str(soup.find('tr' , {'class' : "first"})) #find는 가장 첫번째 부분을 찾는다 = 가장 최신 날짜
        # print(test)
        new_list = test.split('</td>') #최신 날짜를 태그값을 기준으로 리스트에 나누어 저장
        # print(new_list)
        currency_value = [] # 최신 날짜 환율 받을 빈 리스트 생성
        for i in range(len(new_list)):
            temp = new_list[i].split() #  temp 리스트에 공백을 기준으로 나누어 저장
            # print(temp)
            currency_value.append(temp[-1]) #리스트 마지막에 있는 값이 환율 데이터 값
        del currency_value[-1] #리스트 마지막 공백,해쉬태그 값 삭제
        del currency_value[0] # 오늘 날짜 제거

        # print(b)

        currency_key = ['달러', '유로', '엔화', '홍콩달러', 'GBP', '링깃', '루블', '호주달러', '캐나다달러', '뉴질랜드', '싱가포르', '스위스프랑', '랜드(남아프리카)' ,'원화', '디르함(아랍에미리트)',
                        '리얄(카타르)', '포린트(헝가리)',  '즈워티(폴란드)', '크로네(덴마크)', '크로나(스웨덴)',  '크로네(노르웨이)',  '리라(튀르키예즈)',  '페소(필리핀)', '밧(태국)']

        currency_dic = [(currency_key[i], Decimal(currency_value[i])/100) for i in range(len(currency_key))] #딕셔너리 형태로 저장 환율값은 decimal로 정확한 값으로 저장
        currency_dic = dict(currency_dic)
        print(currency_dic)
        return currency_dic

    else:
        return None

```

```py
#stp 환율데이터
import re
import requests
import time
from bs4 import BeautifulSoup
import lxml


def STP_currency():
    html_main_page = requests.get("https://www.bnm.gov.my/exchange-rates/", timeout= 20)
    if html_main_page:
        soup = BeautifulSoup(html_main_page.text, "lxml")

        test = str(soup.select('td')) #문자열로 변환
        test = re.sub('(<([^>]+)>)', '', test) #태그 제거
        test = test.replace('[','')
        test = test.replace(']','')  ## [] 제거
        new_list = test.split(',') #문자열을 콤마를 기준으로 다시 리스트로 저장

        while ' \xa0' in new_list: ## 공백 리스트 제거
            new_list.remove(' \xa0')

        for i in range(len(new_list)): #리스트안 숫자 왼쪽에 공백이 하나있어 매치가 안되기 때문에 왼쪽 공백 제거 ex) ' 230.01'
            new_list[i]= new_list[i].lstrip()

        # print(new_list) # 모든 데이터가 담긴 리스트
        # print(new_list[0])
        index_num = [i for i in range(len(new_list)) if new_list[i] == new_list[0]] #해당 월 1일 인덱스 값 위치
        # print(index_num)
        ### 웹 페이지에 존재하는 환율 데이터 정보를 위해서 인덱스 값을 기준으로 리스트에 저장#####

        one_list = new_list[index_num[0]:index_num[1]] # USD ~ HKD100
        one_list = one_list[-9:] #단락의 최신날짜 환율
        # print(one_list)
        two_list = new_list[index_num[1]:index_num[2]] # THB100 ~ BND
        two_list = two_list[-9:] #단락의 최신날짜 환율
        # print(two_list)
        thr_list = new_list[index_num[2]:] #VND100 ~ EGP
        thr_list = thr_list[-9:] #단락의 최신날짜 환율
        # print(thr_list)

        currency_value = one_list + two_list + thr_list
        # print(currency_value)
        currency_key = ['USD', 'GBP', 'EUR', 'JPY100', 'CHF', 'AUD', 'CAD', 'SGD', 'HKD100',
                        'THB100', 'PHP100', 'TWD100', 'KRW100', 'IDR100', 'SAR100', 'SDR', 'CNY', 'BND',
                        'VND100', 'KHR100', 'NZD', 'MMK100', 'INR100', 'AED100', 'PKR100', 'NPR100', 'EGP']

        currency_dic = [(currency_key[i], currency_value[i]) for i in range(len(currency_key))]
        currency_dic = dict(currency_dic)

        return currency_dic

    else:
        return None
```

## 2.cx_oracle를 이용한 oracle db 연동 및 데이터 insert

```py
import cx_Oracle
import datetime
from STP_currency_rate import STP_currency
from STX_currency_rate import STX_currency
import os
import lxml
from outlook_mail import outlook_mail_db
from outlook_mail import outlook_mail_web

###### 환율데이터 가져오기
stp = STP_currency()
# print(stp)
stx = STX_currency()
# print(stx)
### 오라클 DB 접속
con = cx_Oracle.connect(user = 'apps', password = 'Simmappdev1', dsn='10.101.1.172:1531/MDEV2', encoding = 'UTF-8')
cursor = con.cursor()

################ 중국 시안 환율 입력
if stp is not None and stx is not None: ## 환율데이터가 정상적으로 가져왔고
    if cursor: #오라클에 접속이 완료되었으면
        print("접속 성공")
        dt_now = datetime.datetime.now()

        ##중국 시안 환율 INSERT
        query = [('CNY', 'USD', dt_now.date(), dt_now.date() , 'Corporate STX', stx[0] ,'I')]
        sql ="INSERT INTO gl_daily_rates_interface_temp (from_currency, to_currency, from_conversion_date, to_conversion_date, user_conversion_type, conversion_rate, mode_flag) VALUES (:1,:2,:3,:4,:5,:6,:7)"
        cursor.executemany(sql, query)

        query = [('CNY', 'KRW', dt_now.date(), dt_now.date(), 'Corporate STX', stx[1], 'I')]
        sql = "INSERT INTO gl_daily_rates_interface_temp (from_currency, to_currency, from_conversion_date, to_conversion_date, user_conversion_type, conversion_rate, mode_flag) VALUES (:1,:2,:3,:4,:5,:6,:7)"
        cursor.executemany(sql, query)

        ## 말레이시아 한율 INSERT
        query = [('MYR', 'USD', dt_now.date(), dt_now.date(), 'Corporate STP', stp[0], 'I')]
        sql = "INSERT INTO gl_daily_rates_interface_temp (from_currency, to_currency, from_conversion_date, to_conversion_date, user_conversion_type, conversion_rate, mode_flag) VALUES (:1,:2,:3,:4,:5,:6,:7)"
        cursor.executemany(sql, query)

        query = [('MYR', 'KRW', dt_now.date(), dt_now.date(), 'Corporate STP', stp[1], 'I')]
        sql = "INSERT INTO gl_daily_rates_interface_temp (from_currency, to_currency, from_conversion_date, to_conversion_date, user_conversion_type, conversion_rate, mode_flag) VALUES (:1,:2,:3,:4,:5,:6,:7)"
        cursor.executemany(sql, query)
        print("환율 데이터 입력 완료")

        con.commit()

        con.close()
    else:
        print("DB 접속 실패")
        outlook_mail_db() #담당자에게 메일 송부
else:
    outlook_mail_web()
```
## 3. notification
### win32com.client를 이용한 outlook 메일 송부
```py
import win32com.client


def outlook_mail_db(): #to,cc,subject,body

    outlook = win32com.client.Dispatch("Outlook.Application")
    Rxoutlook = outlook.GetNamespace("MAPI")
    send_mail = outlook.CreateItem(0)

    inbox = Rxoutlook.GetDefaultFolder(6)

    send_mail.To = 'dhbae@simmtech.com'
    send_mail.CC = ''
    send_mail.Subject = "환율크롤링 오라클 db 접속 실패 안내 송부"
    send_mail.HTMLBody = "오라클 db 접속 이상" #내용

    send_mail.send

    return None

# to, cc, subject, body = input('메일정보 입력 : ').split(",")
# outlook_mail(to, cc, subject, body)


def outlook_mail_web(): #to,cc,subject,body

    outlook = win32com.client.Dispatch("Outlook.Application")
    Rxoutlook = outlook.GetNamespace("MAPI")
    send_mail = outlook.CreateItem(0)

    inbox = Rxoutlook.GetDefaultFolder(6)

    send_mail.To = 'dhbae@simmtech.com'
    send_mail.CC = ''
    send_mail.Subject = "환율크롤링 웹사이트 접속 오류"
    send_mail.HTMLBody = "웹사이트 접속 실패 " #내용

    send_mail.send

    return None
```
### oracle procedure을 이용한 notification

<center><img src= "https://user-images.githubusercontent.com/106712669/204946721-126f0508-59a2-48b8-8161-12153c644f67.png" width = "500" height = "500"></center>
<center><img src= "https://user-images.githubusercontent.com/106712669/204947044-4049a31e-df8f-4464-8f6d-753c65b366d4.png" width = "500" height = "500"></center>

## 4. Scheduling
### 작업 스케줄러를 통해서 매일 돌아가도록 설정

<center><img src= "https://user-images.githubusercontent.com/106712669/204947678-102b9d81-7eef-42fa-bdf7-57672c9e2538.png" width = "700" height = "200"></center>


