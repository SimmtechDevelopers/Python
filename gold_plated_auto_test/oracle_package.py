from oracle_db import db_conn, initialize_db_connection
import cx_Oracle
import logging
from datetime import datetime


cnn, cursor = initialize_db_connection()
print("db 연결 성공하였습니다.")



def cbst_common_pkg_PutFileValue(file_name, file_value, org, spec_dept):
    try:

        cursor.callproc('cbst_common_pkg.PutFileValue', [file_name, file_value, org, spec_dept])

        current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        logging.info(
            f'[{current_time}] Procedure executed successfully: file_name={file_name}, file_value={file_value}, org={org}, spec_dept={spec_dept}')
        return 'S'
    except cx_Oracle.Error as error:
        current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        logging.error(
            f'[{current_time}] Error executing procedure: {error}, file_name={file_name}, file_value={file_value}, org={org}, spec_dept={spec_dept}')

        return 'F'

# cnn과 cursor 반환 함수 추가
def get_db_connection():
    return cnn, cursor

# 프로시저 호출
# result = cbst_common_pkg_PutFileValue('SSD03435A00-000',1.136, 101, 0)
# print(result)