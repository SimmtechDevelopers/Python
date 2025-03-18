import os
import json
import subprocess
import shutil
import sys

# 네트워크 경로
network_path = r"\\10.101.2.2\gen_db4\output"

def get_config_path(filename):
    if getattr(sys, 'frozen', False):
        # If the application is run as a bundle, the PyInstaller bootloader
        # extends the sys module by a flag frozen=True and sets the app
        # path into variable _MEIPASS'.
        application_path = os.path.dirname(sys.executable)
    else:
        application_path = os.path.dirname(os.path.abspath(__file__))

    return os.path.join(application_path, filename)

# config 파일 경로 설정
config_path = get_config_path('config.json')
print(config_path)
# config 파일 읽기
with open(config_path, 'r') as config_file:
    folder_mapping = json.load(config_file)


def get_folder_file_values():
    # 폴더와 파일 리스트 읽기
    try:
        folder_file_values = []  # 폴더 명, 파일 명, 파일 내용의 값을 저장할 리스트

        folders = os.listdir(network_path)
        print(f"폴더 목록: {folders}")  # 폴더 목록 출력
        for folder in folders:
            folder_path = os.path.join(network_path, folder)
            if os.path.isdir(folder_path):  # 폴더인지 확인
                mapped_values = folder_mapping.get(folder, {"org": "default_org", "dept_type": 0})  # 폴더명에 매핑된 값 가져오기
                org = mapped_values["org"]
                dept_type = mapped_values["dept_type"]
                # print(f"폴더: {folder}, org: {org}, dept_type: {dept_type}")
                files = os.listdir(folder_path)
                print(files)
                for file in files:
                    file_path = os.path.join(folder_path, file)
                    if os.path.isfile(file_path):  # 파일인지 확인
                        with open(file_path, 'r') as f:
                            value = f.read().strip()  # 파일 내용 읽기
                        # print(f"  파일: {file} 값: {value}")
                        # 폴더 명, 파일 명, 파일 내용, org, dept_type 값을 튜플로 저장하고 리스트에 추가
                        folder_file_values.append((file_path,file, value, org, dept_type))

        return folder_file_values

    except Exception as e:
        print(f"에러가 발생했습니다: {e}")
        return []

def move_file_based_on_result(file_path, result):
    # 파일이 있는 폴더 경로를 기준으로 'Completed' 및 'Failed' 폴더 경로 설정
    base_folder = os.path.dirname(file_path)
    completed_folder = os.path.join(base_folder, 'Completed')
    failed_folder = os.path.join(base_folder, 'Failed')

    # 파일 이동 경로 설정
    if result == 'S':
        dest_path = os.path.join(completed_folder, os.path.basename(file_path))
    else:
        dest_path = os.path.join(failed_folder, os.path.basename(file_path))

    # 파일 이동
    shutil.move(file_path, dest_path)
    print(f'File moved to: {dest_path}')
