import win32com.client
# read mail box
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder - in this case,
                                    # the inbox. You can change that number to reference
                                    # any other folder
messages = inbox.Items
for message in messages:
  print("===============================================================================")  
  print(message.SentOnBehalfOfName)
  print(message.CreationTime)
  print(message.Subject)
  print(message.Body)
  print("===============================================================================")  
  
