import cx_Oracle
from abc import ABC, abstractmethod
from datetime import datetime
from STG_currency_rate  import STG_currency
import os
import sys

#현재 작업 디렉터리

if hasattr(sys, 'frozen'):
    # If the script is compiled to an executable with PyInstaller
    current_dir = os.path.dirname(sys.executable)
else:
    # If the script is run as a regular Python script
    current_dir = os.path.dirname(os.path.abspath(__file__))

lib_dir = os.path.join(current_dir, "instantclient_21_8")
print(os.path.join(current_dir,  "instantclient_21_8"))
print(lib_dir)

## 당일일자 전역변수
now_date = datetime.now().date()

class ExchangeRateProvider(ABC):
    @abstractmethod
    def get_exchange_rates(self) -> list:
        pass

class Db_Connect(ABC):
    @abstractmethod
    def db_connect(self):
        pass

class Database(ABC):
    @abstractmethod
    def insert_exchange_rates(self, rates: tuple):
        pass

#interface에 있는 오늘 날짜 환율 정보 삭제 추상메소드 정의
class toaday_rate_delete(ABC):
    @abstractmethod
    def today_rate_delete(self):
        pass


class StgTodayRateDelete(toaday_rate_delete):

    def __init__(self, db_connect_instance):
        self.db_connect_instance = db_connect_instance

    def today_rate_delete(self):

        con, cursor = self.db_connect_instance.db_connect()

        if cursor is None or con is None:  # 수정된 부분: cursor와 con이 모두 None인 경우를 체크
            print("DB에 연결할 수 없습니다.")
            return

        try:
            # SQL 문 실행
            sql = "DELETE FROM gl_daily_rates_interface WHERE from_conversion_date = :from_conversion_date and user_conversion_type = 'Corporate STG'"
            cursor.execute(sql, {'from_conversion_date': now_date.today()})
            con.commit()
            print('금일 stg 환율 삭제 완료되었습니다.')
        except Exception as e:
            print(f"삭제 중 오류가 발생했습니다: {e}")
        finally:
            cursor.close()
            con.close()



class STGCurrencyRateProvider(ExchangeRateProvider):
    def get_exchange_rates(self) -> list:
        exchange_rates = STG_currency()
        exchange_rates_list = [{'currency': currency, 'rate': rate} for currency, rate in exchange_rates.items()]
        return exchange_rates_list, 'STG', 'JPY'


class ApDbConnect(Db_Connect):


    def db_connect(self):
        try:
            dsn = '''(DESCRIPTION=
                                    (ADDRESS=(PROTOCOL=tcp)(HOST=10.101.1.200)(PORT=1521))
                                (CONNECT_DATA=
                                    (SERVICE_NAME=STERP)
                                    (INSTANCE_NAME=STERP1)
                                )

                            )'''
            cx_Oracle.init_oracle_client(
                # lib_dir=r"C:\Applications\OneSpirit Enterprise Framework\Windows Service Release\CurrencyRateCrawling\instantclient-basic-windows.x64-21.8.0.0.0dbru\instantclient_21_8")
                 lib_dir = os.path.join(current_dir, "instantclient_21_8"))
            self.con = cx_Oracle.connect(user='apps', password='K9k2dic5ua', dsn=dsn, encoding='UTF-8')
            self._cursor = self.con.cursor()
            return self.con, self._cursor  # 성공적으로 연결되었음을 반환
        except cx_Oracle.Error as error:
            print("DB 연결 오류:", error)
            return False  # 연결 실패를 반환

class BatchInsertDatabase(Database):

    def __init__(self, db_connect_instance):
        self.db_connect_instance = db_connect_instance

    # 금일 날짜만 환율을넣는 것이 아니라 튜플로 넘겨받은 date의 모든 환율을 db에 insert 해준다.
    def insert_exchange_rates(self, exchange_rates_data: tuple, org, tp_currency):
        exchange_rates = exchange_rates_data

        con, cursor = self.db_connect_instance.db_connect()

        if cursor is None or con is None:  # 수정된 부분: cursor와 con이 모두 None인 경우를 체크
            print("DB에 연결할 수 없습니다.")
            return
        # 환율정보 insert 시작
        for rate in exchange_rates:
            date = rate['date']
            currency = rate['currency']
            rate_value = rate['rate']

            # 날짜 형식을 datetime 객체로 변경
            conversion_date = datetime.strptime(rate['date'], "%Y-%m-%d")

            try:
                query = [
                    (currency, tp_currency, conversion_date, conversion_date, 'Corporate ' + org, rate_value, 'I')]
                sql = "INSERT INTO gl_daily_rates_interface (from_currency, to_currency, from_conversion_date, to_conversion_date, user_conversion_type, conversion_rate, mode_flag) VALUES (:1,:2,:3,:4,:5,:6,:7)"
                cursor.executemany(sql, query)
                con.commit()
                print(f"{currency} 환율 정보가 성공적으로 삽입되었습니다.")

            except cx_Oracle.Error as error:
                print(f"{currency} 환율 정보 삽입 중 오류 발생:", error)
                con.close()
        con.close()

class DevDbConnect(Db_Connect):
    def __init__(self):
        super().__init__()  # Db_Connect 클래스의 생성자 호출
        self.cursor =None
        OracleClientInitializer.init_oracle_client() # Oracle 클라이언트 라이브러리 초기화


    def db_connect(self):
        try:
            dsn = '''
                      (DESCRIPTION =
                        (ADDRESS = (PROTOCOL = TCP)(HOST = 10.101.1.171)(PORT = 1531))
                            (CONNECT_DATA =
                                (SID = MDEV1)
                        )
                      ) '''
            self._con = cx_Oracle.connect(user='apps', password='Simmappdev1', dsn=dsn, encoding='UTF-8')
            self._cursor = self._con.cursor()

            return self._con,self._cursor  # 성공적으로 연결되었음을 반환
        except cx_Oracle.Error as error:
            print("DB 연결 오류:", error)
            return None,None  # 연결 실패를 반환

class OracleClientInitializer:
    _initialized = False

    @classmethod
    def init_oracle_client(cls):
        if not cls._initialized:
            # cx_Oracle.init_oracle_client(lib_dir=r"C:\Users\simmtech\PycharmProjects\pythonProject\instantclient_21_8")
            cx_Oracle.init_oracle_client(lib_dir = lib_dir)
            cls._initialized =True

class InsertDatabase(Database):

    def __init__(self, db_connect_instance):
        self.db_connect_instance = db_connect_instance

    def insert_exchange_rates(self, exchange_rates_data: tuple):
        exchange_rates, org, tp_currency = exchange_rates_data

        # DB 연결
        con, cursor = self.db_connect_instance.db_connect()
        print(cursor)
        if cursor is None:
            print("DB에 연결할 수 없습니다.")
            return

        # 환율 정보를 받아와서 데이터베이스에 삽입
        for rate in exchange_rates:
            # print(rate['currency'])
            currency = rate['currency']
            # print(currency, type(currency))
            rate_value = rate['rate']
            # print(rate_value,type(rate_value))
            # print(now_date.today())
            try:
                # 환율 정보를 데이터베이스에 삽입하는 SQL 문 실행
                query = [(currency, tp_currency, now_date.today(), now_date.today(), 'Corporate ' + org, rate_value,
                          'I')]
                sql = "INSERT INTO gl_daily_rates_interface (from_currency, to_currency, from_conversion_date, to_conversion_date, user_conversion_type, conversion_rate, mode_flag) VALUES (:1,:2,:3,:4,:5,:6,:7)"

                cursor.executemany(sql, query)
                con.commit()
                print(f"{currency} 환율 정보가 성공적으로 삽입되었습니다.")

            except cx_Oracle.Error as error:
                print(f"{currency} 환율 정보 삽입 중 오류 발생:", error)
                con.close()
        con.close()