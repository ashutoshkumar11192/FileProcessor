#!/usr/bin/env python
# coding: utf-8

# In[6]:


import pandas as pd
import datetime
import numpy as np
import xlsxwriter
import numpy as np
import smtplib
import os
from email.message import EmailMessage
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders


# In[7]:


Subject= 'Amazon CP Shipments in Last Mile'
From="himayat@ecomexpress.in"
To="himayat@ecomexpress.in;Avinashbabu@ecomexpress.in;ganeshaa@ecomexpress.in;chandra.d@ecomexpress.in;janjyoti@ecomexpress.in;kalilulla@ecomexpress.in;thirumavalavan@ecomexpress.in;shaji@ecomexpress.in;rinkun@ecomexpress.in;zeeshankazi@ecomexpress.in;harvinders@ecomexpress.in;rajkumart@ecomexpress.in;sapan.a@ecomexpress.in;rakesh1@ecomexpress.in;yeshpal@ecomexpress.in;tanushree.s@ecomexpress.in;lalitmohanb@ecomexpress.in;bhojp@ecomexpress.in;prakash@ecomexpress.in;ramesh.jaiswar@ecomexpress.in;sanjaym@ecomexpress.in;pawankn@ecomexpress.in;yogeshk@ecomexpress.in;sudhansus@ecomexpress.in;suneetk@ecomexpress.in;kalilulla@ecomexpress.in@ecomexpress.in;thirumavalavan@ecomexpress.in;ramesh.babu@ecomexpress.in;suneetk@ecomexpress.in;chandra.d@ecomexpress.in;ramakrishna@ecomexpress.in;vikash@ecomexpress.in;vivek.sharma@ecomexpress.in;rajeevy@ecomexpress.in;pawankn@ecomexpress.in;sudhansus@ecomexpress.in;anupm@ecomexpress.in;yeshpal@ecomexpress.in;anvas@ecomexpress.in;vishalgupta@ecomexpress.in;abyson@ecomexpress.in;trevor.r@ecomexpress.in;bilala@ecomexpress.in;nithinv@ecomexpress.in;kaluram@ecomexpress.in;apalanikumar@ecomexpress.in;m.lokanathan@ecomexpress.in;anvas@ecomexpress.in;nareshgarg@ecomexpress.in;platinumteam@ecomexpress.in;goh@ecomexpress.in;kaluram@ecomexpress.in;yogeshr@ecomexpress.in;nitink@ecomexpress.in;bilala@ecomexpress.in;apalanikumar@ecomexpress.in;chandra.d@ecomexpress.in;janjyoti@ecomexpress.in;sridharyadav@ecomexpress.in;karnekota.m@ecomexpress.in;m.kumar@ecomexpress.in;perangalla@ecomexpress.in;subhashg@ecomexpress.in;apalanikumar@ecomexpress.in;anoobn@ecomexpress.in;rajeswararao@ecomexpress.in;k.dhiraj@ecomexpress.in;chandrashekara@ecomexpress.in;LM_IHQ@ecomexpress.in;rajsurajsingh@ecomexpress.in;sanayairang@ecomexpress.in;dineshsingh@ecomexpress.in;shoubamd@ecomexpress.in"
Cc="abhilasha.a@ecomexpress.in;adiba@ecomexpress.in;ankit.bansal@ecomexpress.in;Anurag.r@ecomexpress.in;b.krishnananda@ecomexpress.in;CONTROLTOWER@ecomexpress.in;dhirendra.s@ecomexpress.in;dhyani.ashish@ecomexpress.in;harishj@ecomexpress.in;himanshoo.t@ecomexpress.in;kapilgupta@ecomexpress.in;karunesh@ecomexpress.in;mukul.s@ecomexpress.in;nitashaa@ecomexpress.in;pankaj@ecomexpress.in;pankaj.k@ecomexpress.in;prajeetd@ecomexpress.in;radha.n@ecomexpress.in;rahularora@ecomexpress.in;rahul.tandon@ecomexpress.in;rajat@ecomexpress.in;sanjibs@ecomexpress.in;Saurav.kumar@ecomexpress.in;udit.g@ecomexpress.in;kumar.yogesh@ecomexpress.in;nitashaa@ecomexpress.in;b.krishnananda@ecomexpress.in;rakeshl@ecomexpress.in;cloyed.s@ecomexpress.in;sandeep.purwar@ecomexpress.in;jayantb@ecomexpress.in;sanjay.khanna@ecomexpress.in;dushyantsingh@ecomexpress.in;rohitshukla@ecomexpress.in"


# In[8]:


msg =MIMEMultipart()
msg['Subject'] =Subject 
msg['From'] = From
msg['To']= To
msg['Cc'] = Cc


# In[9]:


df1=pd.read_csv("threeday2 .csv",usecols=["airwaybill_number","Destination_DC","Destination_State","CP","RCA","Type.2"])


# In[ ]:


df2=df1[(df1.RCA=='At Origin') | (df1.RCA =='Mid Mile') | (df1.RCA=='At Destination Hub') | (df1.RCA =='At Dc')  | (df1.RCA=='In-transit to LM Dc')]


# In[ ]:


today=datetime.datetime.today().strftime('%Y-%m-%d')
a=datetime.datetime.strptime(today,'%Y-%m-%d')
b = a + datetime.timedelta(days=4)
d4=b.strftime('%Y-%m-%d')
print(today)
print(d4)


# In[ ]:


mask = (df2['CP'] >= today) & (df2['CP'] <= d4)
final=df2.loc[mask]


# In[ ]:


table1=pd.pivot_table(final,values="airwaybill_number",index=["Destination_State"],columns=['RCA'],aggfunc='count',fill_value=0,margins=True,margins_name="Grand Total")


# In[ ]:


table1=table1.assign(sortkey=table1.index == 'Grand Total')                .sort_values(['sortkey','Grand Total'], ascending=[True, False])                .drop('sortkey', axis=1)


# In[ ]:


mask2=(df2['CP']==today)
final2=df2.loc[mask2]


# In[ ]:


table2=pd.pivot_table(final2,values="airwaybill_number",index=["Destination_State"],columns=['RCA'],aggfunc='count',fill_value=0,margins=True,margins_name="Grand Total")


# In[ ]:


table2=table2.assign(sortkey=table2.index == 'Grand Total')                .sort_values(['sortkey','Grand Total'], ascending=[True, False])                .drop('sortkey', axis=1)


# In[ ]:


table2["Destination State"]=table2.index
cols=table2.columns.tolist()
cols = cols[-1:] + cols[:-1]
table2=table2[cols]


# In[ ]:


table1["Destination State"]=table1.index
cols=table1.columns.tolist()
cols = cols[-1:] + cols[:-1]
table1=table1[cols]


# In[ ]:


writer = pd.ExcelWriter('Amazon Pending.xlsx'.format(today), engine='xlsxwriter')


# In[ ]:


final.to_excel(writer, sheet_name='Data',index=False)
table2.to_excel(writer,sheet_name="Today's Pendency",index=False)
table1.to_excel(writer,sheet_name="Overall Pendency",index=False)


# In[ ]:


workbook=writer.book


# In[ ]:


worksheet=writer.sheets["Data"]


# In[ ]:


cell_format=workbook.add_format()
cell_format.set_bold()
cell_format.set_border()
cell_format.set_center_across()
worksheet.set_column("A:F",None,cell_format)


# In[ ]:


for column in final:
    column_width = max(final[column].astype(str).map(len).max(), len(column))
    col_idx = final.columns.get_loc(column)
    writer.sheets['Data'].set_column(col_idx, col_idx, column_width)


# In[ ]:


header_format = workbook.add_format({
    'bold': True,
    'text_wrap': False,
    'valign': 'top',
    'fg_color': '#002060',
    'border': 1 ,
    'center_across':True,
	'font_color':'#FFFFFF'})


# In[ ]:


for col_num, value in enumerate(final.columns.values):
    worksheet.write(0, col_num, value, header_format)


# In[ ]:


worksheet=writer.sheets["Today's Pendency"]
cell_format=workbook.add_format()
cell_format.set_bold()
cell_format.set_border()
cell_format.set_center_across()
worksheet.set_column("A:G",None,cell_format)


# In[ ]:


for column in table2:
    column_width = max(table2[column].astype(str).map(len).max(), len(column))
    col_idx = table2.columns.get_loc(column)
    writer.sheets["Today's Pendency"].set_column(col_idx, col_idx, column_width)


# In[ ]:


header_format = workbook.add_format({
    'bold': True,
    'text_wrap': False,
    'valign': 'top',
    'fg_color': '#002060',
    'border': 1 ,
    'center_across':True,
	'font_color':'#FFFFFF'})


# In[ ]:


for col_num, value in enumerate(table2.columns.values):
    worksheet.write(0, col_num, value, header_format)


# In[ ]:


worksheet=writer.sheets["Overall Pendency"]
cell_format=workbook.add_format()
cell_format.set_bold()
cell_format.set_border()
cell_format.set_center_across()
worksheet.set_column("A:G",None,cell_format)


# In[ ]:


for column in table1:
    column_width = max(table1[column].astype(str).map(len).max(), len(column))
    col_idx = table1.columns.get_loc(column)
    writer.sheets["Overall Pendency"].set_column(col_idx, col_idx, column_width)


# In[ ]:


header_format = workbook.add_format({
    'bold': True,
    'text_wrap': False,
    'valign': 'top',
    'fg_color': '#002060',
    'border': 1 ,
    'center_across':True,
	'font_color':'#FFFFFF'})


# In[ ]:


for col_num, value in enumerate(table1.columns.values):
    worksheet.write(0, col_num, value, header_format)


# In[ ]:


writer.close()


# In[ ]:


df_test=pd.read_excel("Amazon Pending.xlsx",sheet_name=1)


# In[ ]:


def table_style1(x):
    f = x.style.hide_index().set_properties(**{'text-align': 'center',
                                               'border-color':'black',
                                               'border-style':'solid',
                                               'border-width':'thin',
                                               "font-size": "9pt",
                                               'border-collapse':'collapse',
                                               'background':'#f7f7fa',
                                              }).set_table_styles([{'selector':'th','props': [('background', '#5B9BD5'),
                                                                                                     ('text-align', 'center'),
                                                                                                     ('border-color', 'black'),
                                                                                                     ('border-style','solid'),
                                                                                                     ("font-size", "10pt"),
                                                                                                     ("color",'#FFFFFF'),
                                                                                                     ('border-width','thin'),
                                                                                                     ('font-family',['Impact', 'Charcoal', 'sans-serif'])]}])
    return f


# In[ ]:


df_test=table_style1(df_test)


# In[ ]:


html = """    <html>
      <head></head>
      <body>
        <p>Dear Sir/Maam,
        <br>
        These shipments of Amazon CP are in Last Mile and pending for delivery,
        <br>
        Please help to attempt all.
        <br>
        <br>
        {}
        <br>
        <br>
        Regards,<br>
        Himayat Khan
        </p>
      </body>
    </html>
""".format(df_test.to_html())
msg.attach(MIMEText(html,'html'))


# In[ ]:


part = MIMEBase('application', "octet-stream")
part.set_payload(open("Amazon Pending.xlsx", "rb").read())
encoders.encode_base64(part)


# In[ ]:


part.add_header('Content-Disposition', 'attachment; filename="CP {}.xlsx"'.format(today.split('\\')[-1]))
msg.attach(part)


# In[ ]:


smtpObj = smtplib.SMTP('smtp.office365.com:587')
smtpObj.ehlo()
smtpObj.starttls()
smtpObj.login("mail", "Passwork")
try:
    smtpObj.sendmail(From,To.split(";")+Cc.split(";"),msg.as_string())
except Exception as e:
    print("not sent")


# In[ ]:


print("all done")


# In[ ]: