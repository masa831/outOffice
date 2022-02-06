import win32com.client

# outlookオブジェクト作成
outlook = win32com.client.Dispatch("Outlook.Application")
# outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

inbox = outlook.GetDefaultFolder(6)
# outlookのフォルダをすべて格納　
# 第一階層はアカウント　第二階層はアカウントに紐づくフォルダすべて
folders = inbox.Folders
devopsFolder = folders('データ読み込みフォルダ')
messages = devopsFolder.Items

print(inbox)
print(folders)
print(devopsFolder)
print(messages)


def createMail():
    # mail = outlook.CreateItem(0)
    # mail.to = "************@xxx.com"
    # mail.subject = "勤務開始メール（" + str(today) + ")"
    # mail.bodyFormat = 2
    # mail.body = "○○さん"+\
    #             "\n"+\
    #             "\n"+\
    #             "おはようございます。"+"\n"+\
    #             str(start_time)+"より本日の業務を開始します。"+"\n"+\
    #             "よろしくお願いします。"
    
    # mail.display(True)
 
#mail.Send()


