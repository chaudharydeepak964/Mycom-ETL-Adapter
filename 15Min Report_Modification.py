print ( "Ceragon 15Min Report Modification " )

import pandas as pd
import numpy as np
import paramiko 
import datetime
from datetime import datetime, timedelta
import win32com.client as wincl
from openpyxl import load_workbook
import os
import glob
import warnings
import re
warnings.filterwarnings("ignore")

# Auto date


do=(datetime.now() - timedelta(1)).strftime('%d_%m_%Y')
di=(datetime.now() - timedelta(1)).strftime('%Y%m%d')
dm=(datetime.now() - timedelta(1)).strftime('%m-%y')
dn=(datetime.now() - timedelta(0)).strftime('%d-%m-%Y')
dnn=(datetime.now() - timedelta(1)).strftime('%d.%m.%Y')
dnnn=(datetime.now() - timedelta(1)).strftime('%Y-%m-%d')
DA = (datetime.now() - timedelta(0)).strftime('%d-%b-%y')



DA0= datetime.now() - timedelta(0)
DA1= datetime.now() - timedelta(1)
DA2= datetime.now() - timedelta(2)
DA3 = datetime.now() - timedelta(3)
DA4 = datetime.now() - timedelta(4)
DA5= datetime.now() - timedelta(5)
DA6= datetime.now() - timedelta(6)
DA7= datetime.now() - timedelta(7)


dnnn0= datetime.now() - timedelta(0)
dnnn1= datetime.now() - timedelta(1)
dnnn2= datetime.now() - timedelta(2)
dnnn3 = datetime.now() - timedelta(3)
dnnn4 = datetime.now() - timedelta(4)
dnnn5= datetime.now() - timedelta(5)
dnnn6= datetime.now() - timedelta(6)
dnnn7= datetime.now() - timedelta(7)

da7= datetime.now() - timedelta(1)
da6= datetime.now() - timedelta(2)
da5 = datetime.now() - timedelta(3)
da4 = datetime.now() - timedelta(4)
da3 = datetime.now() - timedelta(5)
da2 = datetime.now() - timedelta(6)
da1 = datetime.now() - timedelta(7)

do7= datetime.now() - timedelta(1)
do6= datetime.now() - timedelta(2)
do5 = datetime.now() - timedelta(3)
do4 = datetime.now() - timedelta(4)
do3 = datetime.now() - timedelta(5)
do2 = datetime.now() - timedelta(6)
do1 = datetime.now() - timedelta(7)


dn7= datetime.now() - timedelta(1)
dn6= datetime.now() - timedelta(2)
dn5 = datetime.now() - timedelta(3)
dn4 = datetime.now() - timedelta(4)
dn3 = datetime.now() - timedelta(5)
dn2 = datetime.now() - timedelta(6)
dn1 = datetime.now() - timedelta(7)


dnn7= datetime.now() - timedelta(1)
dnn6= datetime.now() - timedelta(2)
dnn5 = datetime.now() - timedelta(3)
dnn4 = datetime.now() - timedelta(4)
dnn3 = datetime.now() - timedelta(5)
dnn2 = datetime.now() - timedelta(6)
dnn1 = datetime.now() - timedelta(7)

DA7 = datetime.strftime(DA7, '%d-%b-%y')
DA6 = datetime.strftime(DA6, '%d-%b-%y')
DA5 = datetime.strftime(DA5, '%d-%b-%y')
DA4 = datetime.strftime(DA4, '%d-%b-%y')
DA3 = datetime.strftime(DA3, '%d-%b-%y')
DA2 = datetime.strftime(DA2, '%d-%b-%y')
DA1 = datetime.strftime(DA1, '%d-%b-%y')
DA0 = datetime.strftime(DA0, '%d-%b-%y')

da7 = datetime.strftime(da7, '%d_%m_%Y')
da6 = datetime.strftime(da6, '%d_%m_%Y')
da5 = datetime.strftime(da5, '%d_%m_%Y')
da4 = datetime.strftime(da4, '%d_%m_%Y')
da3 = datetime.strftime(da3, '%d_%m_%Y')
da2 = datetime.strftime(da2, '%d_%m_%Y')
da1 = datetime.strftime(da1, '%d_%m_%Y')

do7 = datetime.strftime(do7, '%Y%m%d')
do6 = datetime.strftime(do6, '%Y%m%d')
do5 = datetime.strftime(do5, '%Y%m%d')
do4 = datetime.strftime(do4, '%Y%m%d')
do3 = datetime.strftime(do3, '%Y%m%d')
do2 = datetime.strftime(do2, '%Y%m%d')
do1 = datetime.strftime(do1, '%Y%m%d')

dn7 = datetime.strftime(dn7, '%d-%m-%Y')
dn6 = datetime.strftime(dn6, '%d-%m-%Y')
dn5 = datetime.strftime(dn5, '%d-%m-%Y')
dn4 = datetime.strftime(dn4, '%d-%m-%Y')
dn3 = datetime.strftime(dn3, '%d-%m-%Y')
dn2 = datetime.strftime(dn2, '%d-%m-%Y')
dn1 = datetime.strftime(dn1, '%d-%m-%Y')

dnn7 = datetime.strftime(dnn7, '%d.%m.%Y')
dnn6 = datetime.strftime(dnn6, '%d.%m.%Y')
dnn5 = datetime.strftime(dnn5, '%d.%m.%Y')
dnn4 = datetime.strftime(dnn4, '%d.%m.%Y')
dnn3 = datetime.strftime(dnn3, '%d.%m.%Y')
dnn2 = datetime.strftime(dnn2, '%d.%m.%Y')
dnn1 = datetime.strftime(dnn1, '%d.%m.%Y')

dnnn0 = datetime.strftime(dnnn0, '%Y-%m-%d')
dnnn1 = datetime.strftime(dnnn1, '%Y-%m-%d')
dnnn2 = datetime.strftime(dnnn2, '%Y-%m-%d')
dnnn3 = datetime.strftime(dnnn3, '%Y-%m-%d')
dnnn4 = datetime.strftime(dnnn4, '%Y-%m-%d')
dnnn5 = datetime.strftime(dnnn5, '%Y-%m-%d')
dnnn6 = datetime.strftime(dnnn6, '%Y-%m-%d')

dm7=(datetime.now() - timedelta(1)).strftime('%d_%m_%Y')
dm6=(datetime.now() - timedelta(2)).strftime('%d_%m_%Y')
dm5=(datetime.now() - timedelta(3)).strftime('%d_%m_%Y')
dm4=(datetime.now() - timedelta(4)).strftime('%d_%m_%Y')
dm3=(datetime.now() - timedelta(5)).strftime('%d_%m_%Y')
dm2=(datetime.now() - timedelta(6)).strftime('%d_%m_%Y')
dm1=(datetime.now() - timedelta(7)).strftime('%d_%m_%Y')

da=(datetime.now() - timedelta(1)).strftime('%d_%m_%Y')
do=(datetime.now() - timedelta(1)).strftime('%Y%m%d')
dn=(datetime.now() - timedelta(1)).strftime('%d-%m-%Y')


print(dm7)
print(dm1)
print(da7)


print (" Files Reading Start ")

KCM=pd.read_csv(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Ethernet PM Report Modification\RAW\KAR/Report_'+dm7+'.csv',skiprows=5,encoding= 'unicode_escape')
RCM=pd.read_csv(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Ethernet PM Report Modification\RAW\ROB/Report_'+dm7+'.csv',skiprows=5,encoding= 'unicode_escape')
UCM=pd.read_csv(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Ethernet PM Report Modification\RAW\UPE/Report_'+dm7+'.csv',skiprows=5,encoding= 'unicode_escape')

print (" CM Done, PM Reading Start ")

# KAR - IP-20

KAR=pd.read_csv(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Ethernet PM Report Modification\RAW\KAR/Report_IP-20_15min_'+dm7+'.csv')

print (' KAR Done ')

# ROB- IP -20

r7 = pd.read_csv(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Ethernet PM Report Modification\RAW\ROB/Report_IP-20_15min_' + dm7 + '.csv', encoding='ISO-8859-1')

# ROB- IP -10

r70=pd.read_csv(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Ethernet PM Report Modification\RAW\ROB/Report_IP-10_15min_'+dm7+'.csv')

print (' ROB Done ')


# UPE- IP -20

u7=pd.read_csv(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Ethernet PM Report Modification\RAW\UPE/Report_IP-20_15min_'+dm7+'.csv',encoding= 'unicode_escape')

u70=pd.read_csv(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Ethernet PM Report Modification\RAW\UPE/Report_IP-10_15min_'+dm7+'.csv',encoding= 'unicode_escape')

#UPE=pd.concat([u7,u70])

print (' UPE Done ')

print('All Files Read')



print('Working Start in CM')


UCM=UCM[['Site A Name','Site A Physical Port','Site Z Name','Site Z Physical Port','Link Configuration','Site A IP','Site Z IP']]
KCM=KCM[['Site A Name','Site A Physical Port','Site Z Name','Site Z Physical Port','Link Configuration','Site A IP','Site Z IP']]
RCM=RCM[['Site A Name','Site A Physical Port','Site Z Name','Site Z Physical Port','Link Configuration','Site A IP','Site Z IP']]

CM=pd.concat([UCM,RCM,KCM])

#CM.drop(columns=['Link Configuration','Site A IP','Site Z IP'],inplace=True)

CM1=CM[['Site A Name','Site A Physical Port']]
CM2=CM[['Site Z Name','Site Z Physical Port']]

CM2.rename(columns={'Site Z Name':'Site A Name','Site Z Physical Port':'Site A Physical Port'},inplace=True)

df=pd.concat([CM1,CM2])

df.rename(columns={'Site A Name':'System Name'},inplace=True)

df = df.drop_duplicates(subset=['System Name'])

df['Site A Physical Port'] = df['Site A Physical Port'].str.split('/').str[0]

#Remove last space
df['Site A Physical Port']=df['Site A Physical Port'].str.rstrip()

print('Working start in PM')

# Add Slot For NA from Link Report

KAR=pd.merge(KAR,df,on='System Name',how='left')
KAR['Slot Number'] = KAR['Slot Number'].fillna(KAR['Site A Physical Port'])
KAR['Slot Number'].fillna('NA',inplace=True)
KAR.drop(columns=['Site A Physical Port'],inplace=True)

r7=pd.merge(r7,df,on='System Name',how='left')
r7['Slot Number'] = r7['Slot Number'].fillna(r7['Site A Physical Port'])
r7['Slot Number'].fillna('NA',inplace=True)
r7.drop(columns=['Site A Physical Port'],inplace=True)

r70=pd.merge(r70,df,on='System Name',how='left')
r70['Slot Number'] = r70['Slot Number'].fillna(r70['Site A Physical Port'])
r70['Slot Number'].fillna('NA',inplace=True)
r70.drop(columns=['Site A Physical Port'],inplace=True)


u7=pd.merge(u7,df,on='System Name',how='left')
u7['Slot Number'] = u7['Slot Number'].fillna(u7['Site A Physical Port'])
u7['Slot Number'].fillna('NA',inplace=True)
u7.drop(columns=['Site A Physical Port'],inplace=True)

u70=pd.merge(u70,df,on='System Name',how='left')
u70['Slot Number'] = u70['Slot Number'].fillna(u70['Site A Physical Port'])
u70['Slot Number'].fillna('NA',inplace=True)
u70.drop(columns=['Site A Physical Port'],inplace=True)

## Change data in 'Interface'

KAR['Interface'] = KAR['Interface'].str.replace('E-MC-ABC Group', 'Radio Ethernet', regex=False)
KAR['Interface'] = KAR['Interface'].str.replace('MC-ABC Group', 'Radio Ethernet', regex=False)

r7['Interface'] = r7['Interface'].str.replace('E-MC-ABC Group', 'Radio Ethernet', regex=False)
r7['Interface'] = r7['Interface'].str.replace('MC-ABC Group', 'Radio Ethernet', regex=False)

r70['Interface'] = r70['Interface'].str.replace('E-MC-ABC Group', 'Radio Ethernet', regex=False)
r70['Interface'] = r70['Interface'].str.replace('MC-ABC Group', 'Radio Ethernet', regex=False)

u7['Interface'] = u7['Interface'].str.replace('E-MC-ABC Group', 'Radio Ethernet', regex=False)
u7['Interface'] = u7['Interface'].str.replace('MC-ABC Group', 'Radio Ethernet', regex=False)

u70['Interface'] = u70['Interface'].str.replace('E-MC-ABC Group', 'Radio Ethernet', regex=False)
u70['Interface'] = u70['Interface'].str.replace('MC-ABC Group', 'Radio Ethernet', regex=False)




'''
KAR['Slot Number'].fillna('Slot 1',inplace=True)
r7['Slot Number'].fillna('Slot 1',inplace=True)
u7['Slot Number'].fillna('Slot 1',inplace=True)
'''


with open(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Ethernet PM Report Modification\Output\KAR/Report_IP-20_15min_'+dm7+'.csv', "w", newline='', encoding='utf-8') as f:
    # Write the header without double quotes
    header = ','.join(KAR.columns)
    f.write(header + '\r\n')

    # Write each row with values quoted as needed
    for row in KAR.itertuples(index=False):
        quoted_row = ','.join(
            f'"{str(field)}"' if pd.notna(field) and str(field).strip() != '' else '"N/A"'
            for field in row
        )
        f.write(quoted_row + '\r\n')

print ("KAR Done")


with open(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Ethernet PM Report Modification\Output\ROB/Report_IP-20_15min_' + dm7 + '.csv', "w", newline='', encoding='utf-8') as f:
    # Write the header without double quotes
    header = ','.join(r7.columns)
    f.write(header + '\r\n')

    # Write each row with values quoted as needed
    for row in r7.itertuples(index=False):
        quoted_row = ','.join(
            f'"{str(field)}"' if pd.notna(field) and str(field).strip() != '' else '"N/A"'
            for field in row
        )
        f.write(quoted_row + '\r\n')

print ("ROB 20 Done")

with open(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Ethernet PM Report Modification\Output\ROB/Report_IP-10_15min_' + dm7 + '.csv', "w", newline='', encoding='utf-8') as f:
    # Write the header without double quotes
    header = ','.join(r70.columns)
    f.write(header + '\r\n')

    # Write each row with values quoted as needed
    for row in r70.itertuples(index=False):
        quoted_row = ','.join(
            f'"{str(field)}"' if pd.notna(field) and str(field).strip() != '' else '"N/A"'
            for field in row
        )
        f.write(quoted_row + '\r\n')

print ("ROB 10 Done")

with open(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Ethernet PM Report Modification\Output\UPE/Report_IP-20_15min_'+dm7+'.csv', "w", newline='', encoding='utf-8') as f:
    # Write the header without double quotes
    header = ','.join(u7.columns)
    f.write(header + '\r\n')

    # Write each row with values quoted as needed
    for row in u7.itertuples(index=False):
        quoted_row = ','.join(
            f'"{str(field)}"' if pd.notna(field) and str(field).strip() != '' else '"N/A"'
            for field in row
        )
        f.write(quoted_row + '\r\n')

print ("UPE 20 Done")

with open(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Ethernet PM Report Modification\Output\UPE/Report_IP-10_15min_'+dm7+'.csv', "w", newline='', encoding='utf-8') as f:
    # Write the header without double quotes
    header = ','.join(u70.columns)
    f.write(header + '\r\n')

    # Write each row with values quoted as needed
    for row in u70.itertuples(index=False):
        quoted_row = ','.join(
            f'"{str(field)}"' if pd.notna(field) and str(field).strip() != '' else '"N/A"'
            for field in row
        )
        f.write(quoted_row + '\r\n')

print ("UPE 10 Done")


# Upload on admin --- 

print(" ** Uploading Start on admin ** ")

ssh3=paramiko.SSHClient()
ssh3.set_missing_host_key_policy(paramiko.AutoAddPolicy())
try:
    ssh3.connect(hostname='10.10.10.10',username='admin',password='admin',port=22)
except:
    pass
try:
    ssh3.connect(hostname='11.11.11.11',username='admin',password='admin',port=22)
except:
    pass
sftp_client1=ssh3.open_sftp()



sftp_client1.chdir('/opt/MyLog/TX/Ceragon Server Report/PM/KAR/')
sftp_client1.put(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Ethernet PM Report Modification\Output\KAR/Report_IP-20_15min_'+dm7+'.csv','Report_IP-20_15min_'+dm7+'.csv')

print(" ** KAR Uploaded ** ")

sftp_client1.chdir('/opt/MyLog/TX/Ceragon Server Report/PM/ROB/')
sftp_client1.put(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Ethernet PM Report Modification\Output\ROB/Report_IP-20_15min_' + dm7 + '.csv', 'Report_IP-20_15min_' + dm7 + '.csv')

print(" ** ROB Uploaded ** ")

sftp_client1.chdir('/opt/MyLog/TX/Ceragon Server Report/PM/UPE/')
sftp_client1.put(r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Ethernet PM Report Modification\Output\UPE/Report_IP-20_15min_'+dm7+'.csv','Report_IP-20_15min_'+dm7+'.csv')

print(" ** UPE Uploaded ** ")



sftp_client1.close
ssh3.close

print(" 15 Min Upload done on admin")





