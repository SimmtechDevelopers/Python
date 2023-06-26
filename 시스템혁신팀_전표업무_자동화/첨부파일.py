import time
import pyautogui as pg
import pyperclip as pc

info = [
        r'C:\Users\simmtech\PycharmProjects\pythonProject\maintenance\MES',
        r'C:\Users\simmtech\PycharmProjects\pythonProject\maintenance\EKP',
        r'C:\Users\simmtech\PycharmProjects\pythonProject\maintenance\SHE',
        r'C:\Users\simmtech\PycharmProjects\pythonProject\maintenance\eCIM',
        r'C:\Users\simmtech\PycharmProjects\pythonProject\maintenance\S-EPS',
        r'C:\Users\simmtech\PycharmProjects\pythonProject\license\oracle'
]

info_des = ["기안지", "전자세금계산서","유지보수"]



def scrollbar_move():
    button5location = pg.locateOnScreen('스크롤바.png') # 이미지가 있는 위치를 가져옵니다.
    print(button5location)
    point = pg.center(button5location)
    print(point)
    pg.moveTo(577*1.7,435)
    pg.doubleClick()
    button5location = pg.locateOnScreen('첨부.png')
    print(button5location)
    point = pg.center(button5location)
    print(point)
    pg.moveTo(925,250)
    pg.doubleClick()


def edge_activate():
    window_title = "Edge"
    # 창 가져오기
    window = pg.getWindowsWithTitle(window_title)[0]
    window.activate()
    window.restore()


def copyAndPaste(key):
    pc.copy(key)
    pg.hotkey("ctrl", "v")


def attachments(x):
    pg.press('tab')
    copyAndPaste('E-Accounting Attatchment')
    pg.press('tab')
    copyAndPaste(info_des[x])
    pg.press('tab')
    copyAndPaste(info_des[x])
    pg.press('tab')
    pg.press('tab')


def file_upload_1():

    try:
        time.sleep(1)
        edge_activate()
        button5location = pg.locateOnScreen('찾아보기.png')
        point = pg.center(button5location)
        pg.click(point)
    except:
        pg.press('F5')
        file_upload_1()

def file_path(info,x):
    time.sleep(1)
    pg.press('F4')
    pg.hotkey('ctrl', 'a')
    pg.press('backspace')
    pg.write(info[x])
    pg.press('enter')
    time.sleep(1)

def file_select(x):
    for _ in range(0,4):
        pg.press('tab')
    pg.press('space')
    if x != 0:
        for _ in range(0,x):
            pg.press('down')
    pg.press('enter')
    time.sleep(1)
    pg.press('tab')
    pg.press('enter')
    time.sleep(1)
    pg.hotkey('ctrl', 'w')

def back_to_erp():
    pg.moveTo(100, 10)
    pg.click()
    button5location = pg.locateOnScreen('첨부파일yes.png')
    point = pg.center(button5location)
    pg.click(point)
    pg.moveTo(48,120)
    pg.click()
    pg.press('down')

def next_row():
    pg.press('F4')
    pg.press('down')
    button5location = pg.locateOnScreen('next_None.png')
    point = pg.center(button5location)
    pg.doubleClick(point)

### 파일 첨부 3개 넣는 로직 함수 정의, 인자 받아서 진행(경로)

def file_attach(info,x):
    for i in range(0,3):
        attachments(i)
        time.sleep(2)
        file_upload_1()
        file_path(info,x)
        file_select(i)
        back_to_erp()



###첨부시작###
scrollbar_move()
time.sleep(1)
file_attach(info,0)
for i in range(0,5):
    next_row()
    file_attach(info,i+1)


