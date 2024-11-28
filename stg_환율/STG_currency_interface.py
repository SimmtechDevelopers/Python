import stg_interface_class as stg
import logging
import os

# 로그 파일을 현재 실행되는 스크립트와 같은 디렉토리에서 생성하도록 설정
log_filename = os.path.join(os.path.dirname(__file__), 'exchange_rate.log')

# 로깅 설정
logging.basicConfig(filename=log_filename,
                    level=logging.INFO,
                    format='%(asctime)s:%(levelname)s:%(message)s')


# InsertDatabase를 실행하는 함수를 정의합니다.
def run_insert_exchange_rates():
    # DB 연결 인스턴스 생성 (예: Dev 환경에서 연결)
    # db_connect_instance = stg.DevDbConnect()
    db_connect_instance = stg.ApDbConnect()

    # STGCurrencyRateProvider를 통해 환율 정보를 가져옵니다.
    currency_rate_provider = stg.STGCurrencyRateProvider()
    exchange_rates_data = currency_rate_provider.get_exchange_rates()


    # 각 환율 데이터가 유효한지 확인
    exchange_rates, org, tp_currency = exchange_rates_data
    for rate in exchange_rates:
        if rate['rate'] is None:
            logging.error(f"환율 정보에 NULL 값이 포함되어 있어 작업을 종료합니다.")
            return


    # InsertDatabase 인스턴스를 생성하고, 데이터를 삽입합니다.
    database = stg.InsertDatabase(db_connect_instance)
    database.insert_exchange_rates(exchange_rates_data)


run_insert_exchange_rates()
