import pyautogui
import pyperclip
import time
import csv
from datetime import datetime
from openpyxl import load_workbook

load_wb = load_workbook("C:\\시스템관리팀_전표_자동화.xlsx", data_only=True)  # data_only=True로 해줘야 수식이 아닌 값으로 받아온다.
sheet_name = str(datetime.now().year)[2:] + "년 " + str(datetime.now().strftime('%m')) + "월"  # 예) 23년 01월
load_ws = load_wb[sheet_name]  # 시트 이름으로 불러오기
get_cells = load_ws['J4': "J" + str(load_wb.active.max_row)]  # 엑셀 데이터 행 개수
company, amount, service, gl_date, cut_off, dr_account, tax_code, description, vendor = [], [], [], [], [], [], [], [], []

m365_company_list = ["SEYOUNEC&S"]
communication_company_list = ["KT", "SK BROADBEND", "INPOBANGK", "LG U+", "SK TELECOM", "SYSTEM MANAGEME"]


class InvoiceEntry:
    invoice_condition = ""

    def __init__(self, invoice_condition):
        InvoiceEntry.invoice_condition = invoice_condition

    @staticmethod
    def activate_window():
        win = pyautogui.getWindowsWithTitle('심텍 가동계')[0]
        if win.isMinimized: # 최소화 되어있다면 원복
            win.restore()
        win.activate()  # 윈도우(창) 활성화


class ManipulateKeyboardMouse:
    @staticmethod
    def move_and_click(x, y):
        pyautogui.moveTo(x, y)
        pyautogui.click()

    @staticmethod
    def press_key_number(key, repeat=1):  # 기본값 1로 설정
        for _ in range(repeat):
            pyautogui.keyDown(key)
            pyautogui.keyUp(key)

    @staticmethod
    def copy_and_paste_to_oracle(key):
        pyperclip.copy(key)
        pyautogui.hotkey("ctrl", "v")


class ManipulateData:
    @staticmethod
    def conv_gl_date(date):  # 엑셀 표기 형식의 날짜를 2022-12-30 형태로 바꿔준다.
        if date is None:
            return datetime.today().strftime('%Y-%m-%d')
        else:
            return str(date.year) + "-" + str(date.month) + "-" + str(date.day)


class SetTaxAccountRule:
    @staticmethod
    def tax_code_rule(company_name, service_name):
        # M365 STK 제외, M365의 세금 코드 : C_IN_NONREC_10%
        if service_name in ["M365 STG", "M365 STP", "M365 STX", "M365 T.E TECH"] or "SKT" in service_name:
            return "C_IN_NONREC_10%"
        # 통신 전표 세금 코드 : C2_IN_TAX_10%, 통신 전표 중 오창 : O2_IN_TAX_10%
        elif str(company_name).upper() in communication_company_list:
            if "연수원" in service_name:
                return "C2_IN_TAX_10%"
            else:
                return "O2_IN_TAX_10%" if "오창" in service_name else "C_IN_TAX_10%"
        else:
            return "C_IN_TAX_10%"


class ImportExportData:
    @staticmethod
    def import_data():
        for datas_idx, row in enumerate(get_cells, start=4):
            for _ in row:
                # 청구 금액이 없는 경우 또는 내용이 없는 경우 무시
                if load_ws['D' + str(datas_idx)].value is None or load_ws['B' + str(datas_idx)].value is None:
                    continue
                # m365 에 해당하는 내용만 추가
                if InvoiceEntry.invoice_condition == "m365" and str(load_ws['J' + str(datas_idx)].value).upper() not in m365_company_list:
                    continue
                # 통신 에 해당하는 내용만 추가
                elif InvoiceEntry.invoice_condition == "communication" and str(load_ws['J' + str(datas_idx)].value).upper() not in communication_company_list:
                    continue
                # 기타 에 해당하는 내용만 추가
                elif InvoiceEntry.invoice_condition == "etc" and str(load_ws['J' + str(datas_idx)].value).upper() in communication_company_list + m365_company_list:
                    continue

                if load_ws['A' + str(datas_idx)].value is None:
                    vendor.append(vendor[-1])
                else:
                    vendor.append(load_ws['A' + str(datas_idx)].value)

                company.append(load_ws['J' + str(datas_idx)].value)  # Site Name
                gl_date.append(ManipulateData.conv_gl_date(load_ws['I' + str(datas_idx)].value))  # 품의일자
                amount.append(load_ws['D' + str(datas_idx)].value)  # 청구금액
                service.append(load_ws['B' + str(datas_idx)].value)  # 내용
                cut_off.append(load_ws['G' + str(datas_idx)].value)  # 잡이익
                dr_account.append("1.000.132000." + str(load_ws['K' + str(datas_idx)].value) + ".0000.000000.0.000")  # Dr.Account
                tax_code.append(
                    SetTaxAccountRule.tax_code_rule(load_ws['J' + str(datas_idx)].value, load_ws['B' + str(datas_idx)].value)
                )  # Tax Code

                description.append(
                    str(datetime.now().month) + "월_" + str(vendor[-1]) + "(" + str(load_ws['B' + str(datas_idx)].value) + ")"
                )  # Description 추가

    @staticmethod
    def import_header():
        ManipulateKeyboardMouse.move_and_click(153, 253)  # 첫번째 행 위치 클릭
        for i in range(len(company)):
            if i != 0:
                pyautogui.keyDown('ctrl')
                for j in range(i):
                    pyautogui.press('down')
                ManipulateKeyboardMouse.press_key_number('ctrl')
                ManipulateKeyboardMouse.press_key_number('\t', 2)

            ManipulateKeyboardMouse.press_key_number('s')
            ManipulateKeyboardMouse.press_key_number('\t', 2)

            ManipulateKeyboardMouse.copy_and_paste_to_oracle(company[i].replace("㈜", "(주)"))  # Site Name 입력 (특수 문자 ㈜ 처리)
            ManipulateKeyboardMouse.press_key_number('\t', 2)

            ManipulateKeyboardMouse.copy_and_paste_to_oracle(gl_date[i])  # GL Date 품의 일자 입력

            if InvoiceEntry.invoice_condition == "etc":  # 기타 전표는 Cr.Account가 1.000.000000.210300100.0000.000000.0.000 로 고정되어야 한다.
                ManipulateKeyboardMouse.press_key_number('\t', 6)
                ManipulateKeyboardMouse.copy_and_paste_to_oracle("1.000.000000.210300100.0000.000000.0.000")
                ManipulateKeyboardMouse.press_key_number('\t', 2)
            else:
                ManipulateKeyboardMouse.press_key_number('\t', 8)

            ManipulateKeyboardMouse.copy_and_paste_to_oracle(description[i])
            ManipulateKeyboardMouse.press_key_number('\t', 3)

            ManipulateKeyboardMouse.copy_and_paste_to_oracle('cash')
            ManipulateKeyboardMouse.press_key_number('\t', 4)
            ManipulateKeyboardMouse.press_key_number('v')
            ManipulateKeyboardMouse.press_key_number('\t')
            ManipulateKeyboardMouse.press_key_number('e')

            pyautogui.press('enter')
            if i == len(company) - 1:  # record 추가 시
                pyautogui.hotkey("ctrl", "down")
                ManipulateKeyboardMouse.press_key_number('\t', 2)
        pyautogui.hotkey("ctrl", "s")

    @staticmethod
    def import_cutoff(company_name, cutoff):  # 통신(인포뱅크 제외) 절사액 추가
        cutoff = -cutoff  # 절사액
        if str(company_name).upper() in communication_company_list and str(company_name) != "INPOBANGK":
            ManipulateKeyboardMouse.move_and_click(233, 580)  # 하단 Dr.Account 위치 이동
            time.sleep(1)
            pyautogui.click()
            ManipulateKeyboardMouse.copy_and_paste_to_oracle("1.000.132000.420001700.0000.000000.0.000")
            ManipulateKeyboardMouse.press_key_number('\t', 2)
            ManipulateKeyboardMouse.copy_and_paste_to_oracle(str(cutoff))
            pyautogui.hotkey("ctrl", "s")  # save 버튼

    @staticmethod
    def import_vat():
        items_in_one_scroll = 8
        ManipulateKeyboardMouse.press_key_number("pgup", len(company) // 7 + 2)  # 맨 위로 스크롤하는 횟수

        for i in range(len(company)):
            idx = i + 1
            if idx <= items_in_one_scroll:
                ManipulateKeyboardMouse.move_and_click(354, 252 + (i * 24))
            else:
                ManipulateKeyboardMouse.move_and_click(354, 252 + (7 * 24))
                pyautogui.press("down")

            ManipulateKeyboardMouse.move_and_click(233, 534)
            ManipulateKeyboardMouse.copy_and_paste_to_oracle(dr_account[i])  # Dr.Account 입력
            ManipulateKeyboardMouse.press_key_number('\t', 2)
            ManipulateKeyboardMouse.copy_and_paste_to_oracle(str(amount[i]))  # Amount 입력
            ManipulateKeyboardMouse.press_key_number('\t', 2)
            ManipulateKeyboardMouse.copy_and_paste_to_oracle(tax_code[i])  # Tax Code 입력
            ManipulateKeyboardMouse.press_key_number('\t')
            pyautogui.press("esc")
            pyautogui.hotkey("ctrl", "s")  # save 버튼
            ManipulateKeyboardMouse.move_and_click(124, 695)  # Calculate Tax 버튼
            pyautogui.hotkey("ctrl", "s")  # save 버튼

            if cut_off[i] != 0:
                ImportExportData.import_cutoff(company[i], cut_off[i])  # 절사액 항목 입력

    @staticmethod
    def export_to_csv():
        export_items = []
        for pairs in zip(company, amount, service, gl_date, cut_off, dr_account, tax_code, description):
            export_items.append(list(pairs))
            print(pairs)

        with open("C:\\" + str(datetime.now().month) + "월 시스템관리팀 " + InvoiceEntry.invoice_condition + " 전표 입력 항목.csv", 'w', newline='\n') as f:
            writer = csv.DictWriter(f, fieldnames=["거래처", "청구 금액", "내용", "품의 일자", "잡이익", "Dr.Account", "Tax Code", "Description"])
            writer.writeheader()
            write = csv.writer(f)
            write.writerows(export_items)
