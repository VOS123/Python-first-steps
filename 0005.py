import win32com.client 
object = win32com.client.Dispatch("Outlook.Application")
ns = object.GetNamespace("MAPI")

#list folders met mail

def listfolder(folder):
  for subfolder in folder.Folders:
    try: 
      if len(subfolder.Items) > 0:
        print( len(subfolder.Items) , subfolder.FolderPath  , )
    except:
      continue
    listfolder(subfolder) 
    
for mainfolder in ns.Folders:
   listfolder(mainfolder)

print('Done')
 
