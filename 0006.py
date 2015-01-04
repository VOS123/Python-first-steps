import win32com.client 
import xlsxwriter

object = win32com.client.Dispatch("Outlook.Application")
ns = object.GetNamespace("MAPI")


# Folder name format : \\emailadress  -> use escape chars: \\\\emailadress 
sSearchFolder = "\\\\Jan.de.Vos@unit4.com"

sFile = "Mijninbox2.xlsx"

workbook = xlsxwriter.Workbook(sFile)
format1 = workbook.add_format({'bold': True ,    'fg_color': '#D7E4BC' })

def insert_worksheet(folder):
  row = 0
  col = 0
  if len(folder.Items) > 0 :
    print( ">>>", folder.FolderPath , len(folder.Items) )
    worksheet = workbook.add_worksheet(folder.Name)
    worksheet.set_tab_color('#FF9900')
    worksheet.autofilter('A1:F1')
    worksheet.freeze_panes(1, 1)  
    worksheet.write(row, col,     "Van",               format1)
    worksheet.write(row, col + 1, "Aan",               format1)
    worksheet.write(row, col + 2, "CC",                format1)
    worksheet.write(row, col + 3, "BCC",               format1)
    worksheet.write(row, col + 4, "Ontvangen (datum)", format1)
    worksheet.write(row, col + 5, "Ontvangen (week)",  format1)
    worksheet.write(row, col + 6, "Onderwerp",         format1)
    row += 1
    for message in folder.Items:
      worksheet.write(row, col,     message.SentOnBehalfOfName)
      worksheet.write(row, col + 1, message.To)
      worksheet.write(row, col + 2, message.CC)
      worksheet.write(row, col + 3, message.BCC)
      worksheet.write(row, col + 4, message.CreationTime.strftime("%Y-%m-%d "))
      worksheet.write(row, col + 5, message.CreationTime.strftime("%Y-%W "))
      worksheet.write(row, col + 6, message.Subject)
      row += 1

  
def listfolder(folder):
  #print ( ">>",folder.FolderPath )
  for subfolder in folder.Folders:
    try: 
      if len(subfolder.Items) > 0:
        insert_worksheet(subfolder)
    except:
      continue
    listfolder(subfolder) 
    
for mainfolder in ns.Folders:
   if  mainfolder.FolderPath != sSearchFolder:
     print (">", mainfolder.FolderPath, "Skipped")
   else:
     listfolder(mainfolder)



workbook.set_properties({
    'title':    'Email inbox',
    'subject':  'list incomming mail',
    'author':   'Jan de Vos',
    'company':  'UNIT4',
    'comments': 'Created with Python and XlsxWriter'})
workbook.close() 


print("Done")
  
