import cx_Oracle
import datetime
from STP_currency_rate import STP_currency
from STX_currency_rate import STX_currency
import os
import lxml
from outlook_mail import outlook_mail_db
from outlook_mail import outlook_mail_web
from decimal import Decimal


###### 환율데이터 가져오기
# print(stp)
stx = STX_currency()
# print(stx)
### 오라클 DB 접속
con = cx_Oracle.connect(user = 'apps', password = 'Simmappdev1', dsn='10.101.1.172:1531/MDEV2', encoding = 'UTF-8')
cursor = con.cursor()

################ 중국 시안 환율 입력
if stx is not None: ## 환율데이터가 정상적으로 가져왔고
    if cursor: #오라클에 접속이 완료되었으면
        print("접속 성공")
        dt_now = datetime.datetime.now()

        ##중국 시안 환율 INSERT
        query = [('CNY', 'USD', dt_now.date(), dt_now.date() , 'Corporate STX', stx.get('달러') ,'I')]
        sql ="INSERT INTO gl_daily_rates_interface_temp (from_currency, to_currency, from_conversion_date, to_conversion_date, user_conversion_type, conversion_rate, mode_flag) VALUES (:1,:2,:3,:4,:5,:6,:7)"
        cursor.executemany(sql, query)

        query = [('CNY', 'KRW', dt_now.date(), dt_now.date(), 'Corporate STX', stx.get('원화'), 'I')]
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

################## 말레이시아 환율 입력

