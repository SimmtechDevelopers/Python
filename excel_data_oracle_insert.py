import logging
import pandas as pd
import cx_Oracle
from datetime import datetime
from abc import ABC, abstractmethod
import os
import sys



MAX_ROWS_PER_SHEET = 1048575
# 로그 설정: 실행 시마다 log 파일 초기화 (append 모드 대신 write 모드)
logging.basicConfig(filename='data_insert_log.txt', level=logging.INFO, format='%(asctime)s - %(message)s', filemode='w')

# Oracle 클라이언트 초기화 경로 설정
if hasattr(sys, 'frozen'):
    current_dir = os.path.dirname(sys.executable)  # PyInstaller로 컴파일된 실행 파일인 경우
else:
    current_dir = os.path.dirname(os.path.abspath(__file__))  # 일반 Python 스크립트로 실행되는 경우

lib_dir = os.path.join(current_dir, "instantclient_21_8")


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
            dsn = '''(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=10.101.1.171)(PORT=1531))
                    (CONNECT_DATA=(SID=MDEV1)))'''
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


class MesApDbConnect(Db_Connect):
    def db_connect(self):
        try:
            dsn = '''(DESCRIPTION =
           (ADDRESS = (PROTOCOL = tcp)(HOST=10.101.1.216)(PORT = 1521))
              (CONNECT_DATA =
                 (SID = STMES)
           )
         )
            '''
            cx_Oracle.init_oracle_client(lib_dir=os.path.join(current_dir, "instantclient_21_8"))
            self._con = cx_Oracle.connect(user='mes', password='V78q3kd3nc', dsn=dsn, encoding='UTF-8')
            self._cursor = self._con.cursor()
            return self._con, self._cursor
        except cx_Oracle.Error as error:
            logging.error(f"DB 연결 오류: {error}")
            return None, None

class StpMesApDbConnect(Db_Connect):
    def db_connect(self):
        try:
            dsn = '''(DESCRIPTION=
              (ADDRESS=(PROTOCOL=TCP)(HOST=10.140.1.201)(PORT=1521))
        (CONNECT_DATA =
              (SRVR=DEDICATED)
              (SERVICE_NAME=sustiomes)
        )
      )
            '''
            cx_Oracle.init_oracle_client(lib_dir=os.path.join(current_dir, "instantclient_21_8"))
            self._con = cx_Oracle.connect(user='mes', password='Se3d6ej8ph', dsn=dsn, encoding='UTF-8')
            self._cursor = self._con.cursor()
            return self._con, self._cursor
        except cx_Oracle.Error as error:
            logging.error(f"DB 연결 오류: {error}")
            return None, None

class MesDevDbConnect(Db_Connect):
    def db_connect(self):
        try:
            dsn = '''(DESCRIPTION =
                (ADDRESS = (PROTOCOL = TCP)(HOST = 10.101.1.77)(PORT = 1521))
                (CONNECT_DATA =
                  (SERVICE_NAME=MESDEV2)
                  (INSTANCE_NAME=MESDEV2)
                )
              )
            '''
            cx_Oracle.init_oracle_client(lib_dir=os.path.join(current_dir, "instantclient_21_8"))
            self._con = cx_Oracle.connect(user='mes', password='Simmdev11', dsn=dsn, encoding='UTF-8')
            self._cursor = self._con.cursor()
            return self._con, self._cursor
        except cx_Oracle.Error as error:
            logging.error(f"DB 연결 오류: {error}")
            return None, None





# 데이터 삽입 처리 클래스
class DataInserter:
    def __init__(self, db_connection, config):
        self.db_connection = db_connection
        self.connection, self.cursor = self.db_connection.db_connect()
        self.table_name = config['table_name']
        self.batch_size = int(config['batch_size'])

    def get_column_data_types(self):
        self.cursor.execute(f"SELECT column_name, data_type FROM user_tab_columns WHERE table_name = '{self.table_name.upper()}'")
        columns = self.cursor.fetchall()
        column_types = {col[0]: col[1] for col in columns}
        return column_types

    def insert_data_from_excel(self, excel_path):
        if self.connection is None or self.cursor is None:
            logging.error("DB 연결 실패로 인해 데이터를 삽입할 수 없습니다.")
            return

        start_time = datetime.now()

        try:
            # 엑셀 파일을 여러 시트로 나누어 읽기
            df = pd.read_excel(excel_path, sheet_name=None)  # 모든 시트를 읽어옴
            total_row = 0
            column_types = self.get_column_data_types()  # 컬럼의 데이터 타입 가져오기

            for sheet_name, sheet_df in df.items():
                total_row += sheet_df.shape[0]
                insert_sql = self._generate_insert_sql(sheet_df.columns, column_types)
                insert_values = []

                for index, row in sheet_df.iterrows():
                    row_values = self._get_row_values(row, column_types)
                    if len(row_values) == len(sheet_df.columns):
                        insert_values.append(tuple(row_values))
                    else:
                        logging.error(f"열 개수 불일치: {index + 1}행 데이터 {row_values}")

                    logging.info(f"데이터 삽입 중: {index + 1}/{total_row} 번째 엑셀 로우 삽입, 값: {row_values}")

                    if len(insert_values) >= self.batch_size:
                        self.cursor.executemany(insert_sql, insert_values)
                        self.connection.commit()
                        insert_values.clear()

                # 마지막 배치 데이터 삽입
                if insert_values:
                    self.cursor.executemany(insert_sql, insert_values)
                    self.connection.commit()

        except Exception as e:
            logging.error(f"데이터 삽입 중 오류 발생: {e}")
        finally:
            self.cursor.close()
            self.connection.close()

        end_time = datetime.now()
        elapsed_time = end_time - start_time
        logging.info(f"데이터 삽입 완료. 소요 시간: {elapsed_time}")
        logging.info(f"총 {total_row} row 삽입 완료되었습니다.")
        input("데이터 삽입이 완료되었습니다. Enter 키를 눌러 종료합니다...")

    # def _generate_insert_sql(self, columns, column_types):
    #     return "INSERT INTO {table_name} ({columns}) VALUES ({values})".format(
    #         table_name=self.table_name,
    #         columns=', '.join(columns),
    #         values=', '.join(
    #             [f"TO_DATE(:{i + 1}, 'YYYY-MM-DD HH24:MI:SS')" if column_types.get(col.upper()) == 'DATE' else f':{i + 1}'
    #              for i, col in enumerate(columns)]
    #         )
    #     )
    def _generate_insert_sql(self, columns, column_types):
        # 컬럼명과 값 처리 시, 따옴표와 대소문자에 유의
        columns_str = ', '.join([f'"{col}"' for col in columns])
        values_str = ', '.join(
            [f"TO_DATE(:{i + 1}, 'YYYY-MM-DD HH24:MI:SS')" if column_types.get(col.upper()) == 'DATE' else f':{i + 1}'
             for i, col in enumerate(columns)]
        )

        sql = f"INSERT INTO {self.table_name} ({columns_str}) VALUES ({values_str})"
        logging.info(f"생성된 SQL: {sql}")
        return sql

    def _get_row_values(self, row, column_types):
        row_values = []
        for col, value in row.items():
            if pd.isnull(value):
                row_values.append(None)
            elif column_types.get(col.upper()) == 'NUMBER':
                row_values.append(float(value) if isinstance(value, (float, int)) else None)
            elif column_types.get(col.upper()) == 'DATE':
                if isinstance(value, pd.Timestamp):
                    row_values.append(value.strftime('%Y-%m-%d %H:%M:%S'))
                else:
                    row_values.append(str(value))
            else:
                row_values.append(str(value))
        return row_values

def read_config(config_path):
    config = {}
    with open(config_path, 'r') as f:
        lines = f.readlines()
        for line in lines:
            key, value = line.strip().split('=')
            config[key.strip()] = value.strip()
    return config

# 실행 코드
if __name__ == "__main__":
    config = read_config('config.txt')
    db_type = config.get('db_type', 'Dev').strip()

    if db_type == 'Dev':
        db_conn = DevDbConnect()
    elif db_type == 'Ap':
        db_conn = ApDbConnect()
    elif db_type == 'mesAp':
        db_conn = MesApDbConnect()
    elif db_type == 'StpmesAp':
        db_conn = StpMesApDbConnect()
    elif db_type == 'mesDev':
        db_conn = MesDevDbConnect()
    else:
        print("지원되지 않는 db_type입니다.")
        sys.exit(1)

    data_inserter = DataInserter(db_conn, config)
    data_inserter.insert_data_from_excel('upload_oracle.xlsx')
