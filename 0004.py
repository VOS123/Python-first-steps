import win32com.client as win32
import xlsxwriter

# Read your inbox and write to exccel file
outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")




inbox = outlook.GetDefaultFolder(6)




sFile = "Mijninbox.xlsx"

workbook = xlsxwriter.Workbook(sFile)
format1 = workbook.add_format({'bold': True ,    'fg_color': '#D7E4BC' })

worksheet1 = workbook.add_worksheet("Inbox")
worksheet1.set_tab_color('#FF9900')
worksheet1.autofilter('A1:F1')
worksheet1.freeze_panes(1, 1)



row = 0
col = 0


worksheet1.write(row, col,     "Van",               format1)
worksheet1.write(row, col + 1, "Aan",               format1)
worksheet1.write(row, col + 2, "CC",                format1)
worksheet1.write(row, col + 3, "BCC",               format1)
worksheet1.write(row, col + 4, "Ontvangen (datum)", format1)
worksheet1.write(row, col + 5, "Ontvangen (week)",  format1)
worksheet1.write(row, col + 6, "Onderwerp",         format1)
row += 1

messages = inbox.Items
for message in messages:
  worksheet1.write(row, col,     message.SentOnBehalfOfName)
  worksheet1.write(row, col + 1, message.To)
  worksheet1.write(row, col + 2, message.CC)
  worksheet1.write(row, col + 3, message.BCC)
  worksheet1.write(row, col + 4, message.CreationTime.strftime("%Y-%m-%d "))
  worksheet1.write(row, col + 5, message.CreationTime.strftime("%Y-%W "))
  worksheet1.write(row, col + 6, message.Subject)
  row += 1



workbook.set_properties({
    'title':    'Email inbox' + message.CreationTime.strftime("%Y-%m-%d "),
    'subject':  'list incomming mail',
    'author':   'Jan de Vos',
    'company':  'UNIT4',
    'comments': 'Created with Python and XlsxWriter'})
workbook.close() 

print("Done")
  
