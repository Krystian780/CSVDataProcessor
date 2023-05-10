import zipfile

import pandas as pd
import win32com.client
import os
import datetime
from InboxMessageRetriever import messageRetriever
import re
import matplotlib.pyplot as plt

message = messageRetriever

today_date = str(datetime.date.today())

for message in messages:
 try:
    current_sender = str(message.Sender).lower()
    current_subject = str(message.Subject).lower()
    message_date = str(message.senton.date())
    if re.search('x2',current_subject) != None and message_date == today_date:
      print(current_subject)
      print(current_sender)
      attachments = message.Attachments
      attachment = attachments.Item(1)
      attachment_name = str(attachment).lower()
      attachment.SaveASFile("C:\\Users" + '\\' + attachment_name)
    else:
        pass
    message = messages.GetNext()
 except:
    message = messages.GetNext()

dir_path = r'C:\\Users\\'

res = []
for file in os.listdir(dir_path):
    if file.endswith('.zip'):
        res.append(file)
print(res)

with zipfile.ZipFile("C:\\Users\\" + res[0], 'r') as zip_ref:
    zip_ref.extractall("C:\\Users")

excelFIles = []

for file in os.listdir(dir_path):
        if file.endswith('.xlsx'):
            excelFIles.append(file)

df = pd.read_excel("C:\\Users\\" + excelFIles[0], 'sheet1')
df['CoCode'] = df['Chart Of Accounts'].astype(str).str[:4]
df["CoCode"] = pd.to_numeric(df["CoCode"])
df = pd.pivot_table(df, values=['Req Line #'],
                                index=['CoCode'],
                                aggfunc='count',
                                fill_value=0)

df = df.sort_values(by='Req Line #', ascending=True)

plt.show()
string_name = df.style.set_table_styles([{
              'selector': 'td, th, table'
            , 'props'   : [  ('border', '1px solid lightgrey')
                           , ('border-collapse', 'collapse')
                          ]
            }]).render()

reduced_string = string_name[564:]
firstString = """<style type="text/css">
#T_61f1b td {
  border: 1px solid lightgrey;
  border-collapse: collapse;
}
#T_61f1b  th {
  border: 1px solid lightgrey;
  border-collapse: collapse;
}
#T_61f1b  table {
  border: 1px solid lightgrey;
  border-collapse: collapse;
}
</style>
<table id="T_61f1b">
    <tbody>
    <tr>
      <th id="T_61f1b_level0_row0" class="row_heading level0 row0" >CoCode</th>
      <td id="T_61f1b_row0_col0" class="data row0 col0" >Req Line #</td>
    </tr>
        """

new_string = firstString + reduced_string
print(new_string)



ol=win32com.client.Dispatch("outlook.application")
olmailitem=0x0 #size of the new email
newmail=ol.CreateItem(olmailitem)
newmail.Subject= 'Testing Mail'
newmail.To='krystian.matysek@amadeus.com'
newmail.CC='krystian.matysek@amadeus.com'
newmail.HtmlBody = new_string
# attach='C:\\Users\\admin\\Desktop\\Python\\Sample.xlsx'
# newmail.Attachments.Add(attach)
# To display the mail before sending it
newmail.Display()

writer = pd.ExcelWriter('C:\\Users\\First.xlsx')
df.to_excel(writer, sheet_name='PivotTable')

writer.save()