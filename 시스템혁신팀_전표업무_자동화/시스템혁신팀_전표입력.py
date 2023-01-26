import pyautogui
import pyperclip
import datetime

m = datetime.date.today().month

# 항목별 비용 데이터
info = [
    ['SEYOUNEC&S', 7381250, "MES"],
    ['씨엔씨소프', 2315250, "EKP"],
    ['blue soft', 1500000, "SHE"],
    ['아이엔에쓰', 6900000, "eCIM"],
    ['JS SYSTEM CO.,', 600000, "S-EPS"],  # JS SYSTEM CO.,	2391
    ['인스피언', 400000, "EDI"],
    ['SEYOUNEC&', 19683494, "오라클 라이센스"],
]


def changePrice():
    if m >= 3:  # 3월 이후 부터
        info[0][1] = 7750300
        info[3][1] = 7100000
        if m >= 9:  # 9월 이후 부터
            info[-1][1] = 24725803


def pressTabKey(n):
    for _ in range(n):
        pyautogui.keyDown('\t')
        pyautogui.keyUp('\t')


def pressKey(*keys):
    for key in list(keys):
        pyautogui.keyDown(key)
        pyautogui.keyUp(key)


def copyAndPaste(key):
    pyperclip.copy(key)
    pyautogui.hotkey("ctrl", "v")


def activateWindow():
    win = pyautogui.getWindowsWithTitle('심텍 가동계')[0]
    if win.isMinimized:
        win.restore()
    win.activate()


def moveAndClick(x, y):
    pyautogui.moveTo(x, y)
    pyautogui.click()


# 상단 정보 입력
def insertHeader():
    moveAndClick(153, 253)  # 첫번째 행 위치 클릭
    for i in range(len(info)):
        if i != 0:
            pyautogui.keyDown('ctrl')
            for j in range(i):
                pyautogui.press('down')
            pressKey('ctrl', 'tab', 'tab', 's')
        pressTabKey(2)
        copyAndPaste(info[i][0])
        if info[i][0] == "JS SYSTEM CO.,":
            pressKey('tab', 'esc')
            pressTabKey(9)
        else:
            pressTabKey(10)
        copyAndPaste(str(m) + "월 " + info[i][2] + " 유지보수 건")
        pressTabKey(3)
        copyAndPaste('cash')
        pressTabKey(4)
        pressKey('v', 'tab', 'e', 'enter')
        if i == len(info) - 1:
            pressTabKey(7)


# 하단 부가세 계산
def insertVat():
    for i in range(len(info)):
        moveAndClick(354, 252 + (i * 24))
        moveAndClick(233, 534)
        copyAndPaste("1.000.133010.520028021.0000.000000.0.000")
        pressTabKey(2)
        copyAndPaste(str(info[i][1]))
        pressTabKey(2)
        copyAndPaste("C_IN_TAX_10 ")
        pressKey('\t')
        pyautogui.press("esc")
        pyautogui.hotkey("ctrl", "s")  # save 버튼
        moveAndClick(124, 695)  # Calculate Tax 버튼 클릭
        pyautogui.hotkey("ctrl", "s")  # save 버튼


changePrice()
activateWindow()
insertHeader()
pyautogui.hotkey("ctrl", "s")  # save 버튼
insertVat()
