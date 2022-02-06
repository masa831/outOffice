import win32com.client

# outlookオブジェクト作成　outlook起動
outlook = win32com.client.Dispatch("Outlook.Application")
# outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# 指定したフォルダの情報を取得　引数によってどのフォルダを取得するかを制御する
# 別スクリプトでインデックスとフォルダの関係を把握する→researchFolder.py
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

def loadMailInfo():
    pass

def createMail():
    pass
    # メール新規作成
    # mail = outlook.CreateItem(0)
    # 送付先アドレス
    # mail.to = "************@xxx.com"
    # メールタイトル
    # mail.subject = "勤務開始メール（" + str(today) + ")"
    # mail.bodyFormat = 2
    # メール本文　改行は+\
    # mail.body = "○○さん"+\
    #             "\n"+\
    #             "\n"+\
    #             "おはようございます。"+"\n"+\
    #             str(start_time)+"より本日の業務を開始します。"+"\n"+\
    #             "よろしくお願いします。"
    
    # メールの表示
    # mail.display(True)
    # メール自動送信
    # mail.Send()

def deleteMail():
    pass



