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

