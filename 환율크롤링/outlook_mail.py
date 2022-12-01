import win32com.client


def outlook_mail_db(): #to,cc,subject,body

    outlook = win32com.client.Dispatch("Outlook.Application")
    Rxoutlook = outlook.GetNamespace("MAPI")
    send_mail = outlook.CreateItem(0)

    inbox = Rxoutlook.GetDefaultFolder(6)

    send_mail.To = 'dhbae@simmtech.com'
    send_mail.CC = ''
    send_mail.Subject = "환율크롤링 오라클 db 접속 실패 안내 송부"
    send_mail.HTMLBody = "오라클 db 접속 이상" #내용

    send_mail.send

    return None

# to, cc, subject, body = input('메일정보 입력 : ').split(",")
# outlook_mail(to, cc, subject, body)


def outlook_mail_web(): #to,cc,subject,body

    outlook = win32com.client.Dispatch("Outlook.Application")
    Rxoutlook = outlook.GetNamespace("MAPI")
    send_mail = outlook.CreateItem(0)

    inbox = Rxoutlook.GetDefaultFolder(6)

    send_mail.To = 'dhbae@simmtech.com'
    send_mail.CC = ''
    send_mail.Subject = "환율크롤링 웹사이트 접속 오류"
    send_mail.HTMLBody = "웹사이트 접속 실패 " #내용

    send_mail.send

    return None