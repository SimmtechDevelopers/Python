from oracle_db import db_conn, initialize_db_connection
from oracle_package import cbst_common_pkg_PutFileValue, get_db_connection
from get_data_nas import get_folder_file_values, move_file_based_on_result
import logging
import schedule
import time


# oracle_package.py에서 cnn과 cursor를 받아옵니다.
cnn, cursor = get_db_connection()

folder_file_values = get_folder_file_values()
print(folder_file_values)

# 리스트를 반복문을 통해 프로시저 호출
for file_info in folder_file_values:
    file_path, file_name, file_value, org, spec_dept = file_info  # 튜플 언패킹
    result = cbst_common_pkg_PutFileValue(file_name, file_value, org, spec_dept)
    print(f'File: {file_name}, Result: {result}')

    # 파일 이동
    move_file_based_on_result(file_path, result)

# 모든 작업이 끝난 후에 커서와 연결을 닫습니다.
cursor.close()
cnn.close()
logging.info('Database connection closed')


