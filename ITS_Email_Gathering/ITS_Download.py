import re
import pyodbc
import win32com.client
from io import StringIO
import os


def get_data_from_emails(outlook, target_subject_prefix):
    namespace = outlook.GetNamespace("MAPI")
    inbox = namespace.GetDefaultFolder(6)
    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)

    results = []
    senders = []
    cc_list = []
    process_codes = []
    original_messages = []

    for message in messages:
        try:
            if message.UnRead and message.Subject.startswith(target_subject_prefix):
                received_time = message.ReceivedTime
                if received_time is None:
                    continue

                print(f"메일 수신 시간 (로컬 시간): {received_time}")
                body = message.Body

                # 정규 표현식에서 백슬래시를 올바르게 이스케이프
                body_match = re.search(r'ProcessCode\s*:\s*(\d+)', body)
                process_code = body_match.group(1) if body_match else "140"

                # [Lot]와 [End] 사이의 내용 추출
                match = re.search(r'\[Lot\](.*?)\[End\]', body, re.DOTALL)
                if match:
                    content = match.group(1).strip()
                    data = re.findall(r'([\w\-]+)\s*:\s*(\d+)', content)
                    print(data)

                    valid_data = []
                    for key, value in data:
                        if re.match(r'^[A-Za-z0-9\-]+$', key) and re.match(r'^\d+$', value):
                            valid_data.append((key, value))
                        else:
                            print(f"잘못된 포맷 발견: {key} : {value}")
                            valid_data = []
                            break

                    if valid_data:
                        results.append(valid_data)
                        senders.append(message.SenderEmailAddress)
                        cc_list.append(message.CC)
                        process_codes.append(process_code)
                        original_messages.append(message)

                message.UnRead = False

        except Exception as e:
            print(f"이메일 처리 중 오류 발생: {str(e)}")
            message.UnRead = False
            continue

    return results, senders, cc_list, process_codes, original_messages


def query_data_from_sql(data, process_code):
    # SQL Server 연결 정보 설정
    server = '10.101.1.190'
    database = 'ITS'
    username = 'itsuser'
    password = 'itsuser:'
    driver = '{ODBC Driver 17 for SQL Server}'

    connection_string = f"DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={password}"

    # SQL Server 연결
    connection = pyodbc.connect(connection_string)
    cursor = connection.cursor()
    print("SQL Server 연결 완료")

    lot_values = [item[0] for item in data]  # key값만 추출

    # 결과 저장을 위한 딕셔너리 (LotNumber를 기준으로 그룹화)
    grouped_result = {}

    def rows_to_dict(cursor, rows):
        """Convert pyodbc rows to dictionary format"""
        columns = [column[0] for column in cursor.description]
        return [dict(zip(columns, row)) for row in rows]

    # 첫 번째 쿼리 실행
    for lot_value in lot_values:
        lot_value_trimmed = lot_value[:12].strip()  # '-A' 또는 '-00'을 제거하고, 첫 12자리만 추출 후 공백 제거

        query = f"""
            SELECT lm.LotNumber, lm.*
            FROM dbo.pts_LotMaster lm
            WHERE lm.LotNumber = SUBSTRING('{lot_value_trimmed}', 1, 12) + '-00'
        """
        cursor.execute(query)
        rows = cursor.fetchall()
        rows = rows_to_dict(cursor, rows)
        for row in rows:
            lot_number_trimmed_from_db = row['LotNumber'].strip()  # DB에서 가져온 LotNumber에서 공백 제거
            if lot_number_trimmed_from_db not in grouped_result:
                grouped_result[lot_number_trimmed_from_db] = {"lot_info": [], "strip_info": []}
            grouped_result[lot_number_trimmed_from_db]["lot_info"].append(row)

    # 두 번째 쿼리 실행
    for lot_value in lot_values:
        lot_value_trimmed = lot_value[:12].strip()  # '-A' 또는 '-00'을 제거하고, 첫 12자리만 추출 후 공백 제거

        query_2 = f"""
            SELECT s.StripID, s.PCSCol, s.PCSRow
            FROM dbo.pts_StripBCL2 s
            WHERE s.ProcessCode = {process_code}
                AND s.StripID LIKE SUBSTRING('{lot_value_trimmed}', 1, 12) + '%'
            GROUP BY s.StripID, s.PCSCol, s.PCSRow
            ORDER BY s.StripID, s.PCSCol, s.PCSRow
        """
        cursor.execute(query_2)
        rows = cursor.fetchall()
        rows = rows_to_dict(cursor, rows)
        if rows:
            for row in rows:
                lot_number_trimmed_from_db = row['StripID'][:12].strip()  # DB에서 가져온 StripID에서 LotNumber 추출 후 공백 제거
                # 12자리만 비교
                for grouped_lot_number in grouped_result.keys():
                    if lot_number_trimmed_from_db == grouped_lot_number[:12]:
                        grouped_result[grouped_lot_number]["strip_info"].append(row)
        else:
            print(f"LotNumber {lot_value_trimmed}에 대한 strip 정보가 없습니다.")    # 커넥션 종료
    cursor.close()
    connection.close()

    return grouped_result


def save_results_to_memory(grouped_result, process_codes):
    # 각 LotNumber에 대해 개별 파일 생성
    file_paths = []

    for lot_number, result in grouped_result.items():
        lot_info = result["lot_info"]
        strip_info = result["strip_info"]

        # Lot 정보가 존재하는 경우 처리
        if lot_info:
            lot_record = lot_info[0]  # 첫 번째 항목이 Lot의 상세 정보
            management_code = lot_record['ManagementCode'].strip()  # 뒤의 공백 제거
            lot_number_trimmed = lot_record['LotNumber'].strip()  # 뒤의 공백 제거
        else:
            continue  # Lot 정보가 없다면 건너뛰기

        # 해당 Lot의 Process Code 가져오기
        process_code = process_codes.pop(0) if process_codes else "Unknown"

        # 파일 경로 생성
        file_name = f"{lot_number_trimmed}.txt"
        file_path = os.path.join(os.getcwd(), file_name)
        file_paths.append(file_path)

        # 파일 내용 작성
        with open(file_path, 'w', encoding='utf-8') as file_content:
            file_content.write(f"Management Code : {management_code}\n")
            file_content.write(f"Lot Number : {lot_number_trimmed}\n")
            file_content.write(f"Process Code : {process_code}\n")
            file_content.write(f"Total Count : {len(strip_info)}\n")
            file_content.write("\n")

            for strip in strip_info:
                strip_id = strip['StripID'].strip()  # StripID에서 앞 12자리를 가져옴 (LotNumber 부분)
                pcs_col = strip['PCSCol']
                pcs_row = strip['PCSRow']
                file_content.write(f"{strip_id} {pcs_col},{pcs_row}\n")

            file_content.write("EOL")

    return file_paths


def send_email_with_attachment(original_message, file_paths, to_email, cc_email=None):
    # Outlook 애플리케이션에 연결
    mail = original_message.Reply()  # 회신 생성

    # 메일 제목 및 수신자 설정
    mail.Subject = f"Re: {original_message.Subject}"
    mail.To = to_email

    # 참조자 설정 (있을 경우)
    if cc_email:
        mail.CC = cc_email

    # 메일 본문 설정
    mail.Body = "첨부된 파일에서 요청하신 데이터를 확인하실 수 있습니다."

    # 파일을 메모리에서 읽어와서 첨부
    for file_path in file_paths:
        mail.Attachments.Add(file_path)

    # 메일 전송
    mail.Send()
    print("메일 전송 완료!")

    # 임시 파일 삭제
    for file_path in file_paths:
        os.remove(file_path)
