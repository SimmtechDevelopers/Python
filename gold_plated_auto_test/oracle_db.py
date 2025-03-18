import logging
import cx_Oracle
from datetime import datetime
from abc import ABC, abstractmethod
import os
import sys


# 로그 설정: 실행 시마다 log 파일 초기화 (append 모드 대신 write 모드)
logging.basicConfig(filename='data_insert_log.txt', level=logging.INFO, format='%(asctime)s - %(message)s', filemode='w')

# Oracle 클라이언트 초기화 경로 설정
if hasattr(sys, 'frozen'):
    current_dir = os.path.dirname(sys.executable)  # PyInstaller로 컴파일된 실행 파일인 경우
else:
    current_dir = os.path.dirname(os.path.abspath(__file__))  # 일반 Python 스크립트로 실행되는 경우

lib_dir = os.path.join(current_dir, "instantclient_21_8")

# 현재 코드의 실행 경로 출력
print(f"현재 코드 실행 경로: {current_dir}")
# print(f"Oracle client directory: {lib_dir}")
# logging.info(f"Oracle client directory: {lib_dir}")

# 환경 변수에 경로 추가
os.environ["PATH"] = lib_dir + ";" + os.environ["PATH"]
# print(f"Updated PATH: {os.environ['PATH']}")

# DB 연결 기본 클래스
class Db_Connect(ABC):
    @abstractmethod
    def db_connect(self):
        pass


class OracleClientInitializer:
    _initialized = False

    @classmethod
    def init_oracle_client(cls):
        if not cls._initialized:
            cx_Oracle.init_oracle_client(lib_dir=lib_dir)
            cls._initialized = True


# Dev 환경 DB 연결
class DevDbConnect(Db_Connect):
    def __init__(self):
        super().__init__()
        self.cursor = None
        OracleClientInitializer.init_oracle_client()

    def db_connect(self):
        try:
            dsn = '''(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.101.1.178)(PORT=1531))
                    (CONNECT_DATA=(SID=MDEV3)))'''
            self._con = cx_Oracle.connect(user='apps', password='Simmappdev1', dsn=dsn, encoding='UTF-8')
            self._cursor = self._con.cursor()
            return self._con, self._cursor
        except cx_Oracle.Error as error:
            print("DB 연결 오류:", error)
            return None, None


# Ap 환경 DB 연결
class ApDbConnect(Db_Connect):
    def db_connect(self):
        try:
            dsn = '''(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.101.1.200)(PORT=1521))
                    (CONNECT_DATA=(SERVICE_NAME=STERP)(INSTANCE_NAME=STERP1)))'''
            cx_Oracle.init_oracle_client(lib_dir=os.path.join(current_dir, "instantclient_21_8"))
            self._con = cx_Oracle.connect(user='apps', password='K9k2dic5ua', dsn=dsn, encoding='UTF-8')
            self._cursor = self._con.cursor()
            return self._con, self._cursor
        except cx_Oracle.Error as error:
            print("DB 연결 오류:", error)
            return None, None


# def read_config(config_path):
#     config = {}
#     with open(config_path, 'r') as f:
#         lines = f.readlines()
#         for line in lines:
#             key, value = line.strip().split('=')
#             config[key.strip()] = value.strip()
#     return config

def read_config(filename):
    config = {}
    if getattr(sys, 'frozen', False):
        # If the application is run as a bundle, the PyInstaller bootloader
        # extends the sys module by a flag frozen=True and sets the app
        # path into variable _MEIPASS'.
        application_path = os.path.dirname(sys.executable)
    else:
        application_path = os.path.dirname(os.path.abspath(__file__))

    config_path = os.path.join(application_path, filename)
    with open(config_path, 'r') as f:
        lines = f.readlines()
        for line in lines:
            key, value = line.strip().split('=')
            config[key.strip()] = value.strip()
    return config


db_conn = None

# 실행 코드
def initialize_db_connection():
    global db_conn
    config = read_config('config.txt')
    db_type = config.get('db_type', 'Dev').strip()

    if db_type == 'Dev':
        db_conn = DevDbConnect()
        con, cursor =db_conn.db_connect()
        # print("개발계 접속 완료")
    elif db_type == 'Ap':
        db_conn = ApDbConnect()
        con, cursor = db_conn.db_connect()
    else:
        print("지원되지 않는 db_type입니다.")
        sys.exit(1)
    return con, cursor

# initialize_db_connection()