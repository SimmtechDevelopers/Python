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

print(STX_currency().get('달러'))
