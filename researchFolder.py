import win32com.client
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
for i in range(100):
	try:
		box = outlook.GetDefaultFolder(i)
		name = box.Name
		print(i, name)
	except:
		pass

# 別ver
inbox = outlook.GetDefaultFolder(6)
folders  = inbox.Folders
# devoFolder = folders('読み込みたいフォルダの名前')
for folder in folders:
    print('Name: ' + folder.name)
