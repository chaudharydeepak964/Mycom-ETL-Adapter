print("Ceragon Ethernet 15Min PM  SFTP STARTED")

import pandas as pd
import numpy as np
import paramiko 
import datetime
from datetime import datetime, timedelta
import win32com.client as wincl
from openpyxl import load_workbook
import os
import glob


# Auto date

d1= datetime.now() - timedelta(1)
d2= datetime.now() - timedelta(2)
d3 = datetime.now() - timedelta(3)
d4 = datetime.now() - timedelta(4)
d5 = datetime.now() - timedelta(5)
d6 = datetime.now() - timedelta(6)
d7 = datetime.now() - timedelta(7)

d1 = datetime.strftime(d1, '%d_%m_%Y')
d2 = datetime.strftime(d2, '%d_%m_%Y')
d3 = datetime.strftime(d3, '%d_%m_%Y')
d4 = datetime.strftime(d4, '%d_%m_%Y')
d5 = datetime.strftime(d5, '%d_%m_%Y')
d6 = datetime.strftime(d6, '%d_%m_%Y')
d7 = datetime.strftime(d7, '%d_%m_%Y')


da=(datetime.now() - timedelta(1)).strftime('%d_%m_%Y')
do=(datetime.now() - timedelta(1)).strftime('%Y%m%d')
dm=(datetime.now() - timedelta(1)).strftime('%m-%y')
dn=(datetime.now() - timedelta(0)).strftime('%d-%m-%Y')
dn1=(datetime.now() - timedelta(1)).strftime('%d-%m-%Y')

dmm=(datetime.now() - timedelta(1)).strftime('%m-%y')


path = r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Ethernet PM Report Modification\RAW'

for folder, subfolders,files in os.walk(path):
    for fl in files:
        if fl.endswith(".csv" ):#or fl.endswith('.txt'):
            path = os.path.join(folder, fl)
            os.remove(path)

path = r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Ethernet PM Report Modification\Output'

for folder, subfolders,files in os.walk(path):
    for fl in files:
        if fl.endswith(".csv" ):#or fl.endswith('.txt'):
            path = os.path.join(folder, fl)
            os.remove(path)
            
print("All Files Deleted!")


######################## SFTP at 39 ########################


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

print(" SERVER CONNECTED")

print(" DOWNLOADING Latest Day Full Link Files ")

    

try:
    sftp_client1.chdir('/opt/mycom/cm1/UPE/')
    sftp_client1.get('Report_'+d2+'.csv' ,  r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Ethernet PM Report Modification\RAW\UPE\Report_'+da+'.csv')
    print("4.Full link UPE.csv downloaded")
except Exception as e:
    print(e)


try:
    sftp_client1.chdir('/opt/mycom/cm1/KAR/')
    sftp_client1.get('Report_'+d2+'.csv' ,  r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Ethernet PM Report Modification\RAW\KAR\Report_'+d1+'.csv')
    print("6.KAR downloaded",d1)
except Exception as e:
    print(e)


try:
    sftp_client1.chdir('/opt/mycom/cm1/ROB/')
    sftp_client1.get('Report_'+d2+'.csv' ,  r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Ethernet PM Report Modification\RAW\ROB\Report_'+d1+'.csv')
    print("7.ROB downloaded",d1)
except Exception as e:
    print(e)


   
sftp_client1.close
ssh3.close


## .39 admin ####

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



#################################### UPE Radio PM REPORT ###########################
    

#IP - 20

try:
    sftp_client1.chdir('/opt/mycom/pm1/Test_2/UPE')
    sftp_client1.get('Report_IP-20_15min_'+d1+'.csv' ,  r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Ethernet PM Report Modification\RAW\UPE\Report_IP-20_15min_'+d1+'.csv')
    print("1.UPE Ethernet pm file downloaded",d1)
except Exception as e:
    print(e)


#IP - 10

try:
    sftp_client1.chdir('/opt/mycom/pm1/Test_2/UPE')
    sftp_client1.get('Report_IP-10_15min_'+d1+'.csv' ,  r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Ethernet PM Report Modification\RAW\UPE\Report_IP-10_15min_'+d1+'.csv')
    print("1.UPE Ethernet pm file downloaded",d1)
except Exception as e:
    print(e)


print("7.UPE PM downloaded",d1)



#################################### ROB Radio PM REPORT ###########################

try:
    sftp_client1.chdir('/opt/mycom/pm1/Test_2/ROB')
    sftp_client1.get('Report_IP-20_15min_'+d1+'.csv' ,  r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Ethernet PM Report Modification\RAW\ROB\Report_IP-20_15min_'+d1+'.csv')
    print("1.UPE Ethernet pm file downloaded",d1)
except Exception as e:
    print(e)


try:
    sftp_client1.chdir('/opt/mycom/pm1/Test_2/ROB')
    sftp_client1.get('Report_IP-10_15min_'+d1+'.csv' ,  r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Ethernet PM Report Modification\RAW\ROB\Report_IP-10_15min_'+d1+'.csv')
    print("1.UPE Ethernet pm file downloaded",d1)
except Exception as e:
    print(e)



#################################### KAR Radio PM REPORT ###########################
    
try:
    sftp_client1.chdir('/opt/mycom/pm1/Test_2/KAR')
    sftp_client1.get('Report_IP-20_15min_'+d1+'.csv' ,  r'C:\Users\COR1736664\Desktop\Deepak\ALL CODE\Ceragon Ethernet PM Report Modification\RAW\KAR\Report_IP-20_15min_'+d1+'.csv')
    print("1.KAR Ethernet pm file downloaded",d1)
except Exception as e:
    print(e)


sftp_client1.close
ssh3.close

print(" **All RAW Downloaded",d1)
