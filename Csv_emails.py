# Extraer emails de Outlook a una hoja de Excel

import win32com.client
import csv

msglst = [('Subject', 'Body')]   # initialize the list and set the headers

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6).Folders.Item("Mystatsonline")
messages = inbox.Items

message = messages.GetFirst()

while message:
  msglst.append((message.Subject, message.Body))   # append each subject/body pair to the list
  message = messages.GetNext()

with open('messages_list.csv','w', newline='', encoding='utf-8') as f:
    wrt = csv.writer(f, dialect='excel')
    wrt.writerows(msglst)   # write each subject/body pair as a new line of the csv file