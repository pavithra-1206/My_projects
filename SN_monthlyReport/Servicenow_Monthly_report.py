import numpy as np
import cx_Oracle
import base64
import smtplib
import csv
from email.mime.multipart import MIMEMultipart
from datetime import datetime, timedelta
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email import encoders
from tabulate import tabulate
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
import time
import os

# %matplotlib inline
print('Script execution started.....')

#######################Connecting to DATAMART database #######################
username = base64.b64decode(b'VUhHXzAwMTEyMTk1Ng==').decode('utf-8')
password = base64.b64decode(b'QXByQDIwMjE=').decode('utf-8')
connection = cx_Oracle.connect(username, password, 'ep06.uhc.com:1521/itm0pr01svc.uhc.com')
print('connected to DATAMART database')
cursor = connection.cursor()

# Database Queries
Monthly_tickets = """
SELECT  SIST.IN_ID AS INCIDENT_NUMBER,
        SIS.PRIORITY_CODE AS PRIORITY_CODE,  
        SIS.PROBLEM_STATUS AS STATUS      ,
        SIST.IN_NAME AS ASSGNMENT_GROUP,
        SIST.IN_DATE_FROM AS OPEN_TIME,
        SIST.IN_DATE_TO AS CLOSED_TIME, 
        SIST.IN_SPENTTIME AS SPENT_TIME,
        decode(SIS.UH_ASSIGNEE_FULL_NAME,NULL,'UNASSIGNED',UH_ASSIGNEE_FULL_NAME) as ASSIGNED_TO,
        --SIS.UH_RESTORATION_GOAL as SLA
        DECODE (UH_RESTORATION_GOAL, 0,'Breached SLA',1,'MET SLA') as SLA
FROM SM_DM.SM_INCIDENTS_SPENTTIME SIST,SM_DM.SM_INCIDENTS SIS  
where SIST.IN_ID=SIS.IN_ID
AND IN_NAME IN('UPM (UNT) - OSW','MPE - Member Product Eligibility - SRVC','Consumer360 - SPT','BILLING-RECEIVABLE MGMT SYS - APP','GATEWAY','UHC - OIL API (SUP)')       
AND to_char( sysdate-2, 'mm' ) = to_char( IN_DATE_FROM, 'mm' )  AND TO_CHAR(IN_DATE_FROM, 'yyyy') = TO_CHAR(SYSDATE, 'yyyy')
AND PROBLEM_STATUS != 'Cancelled'
--GROUP BY SIST.IN_ID;
order by ASSGNMENT_GROUP,SLA asc
"""
############## SQL to csv exporting ###################3
cursor.execute(Monthly_tickets)
# path = "//ms/userdata/056/mpavithr/Documents/PYTHON_SCRIPTS"
path = "C:/Users/mpavithr/Documents/UHG/MPE/Scripts/PYTHON SCRIPTS/Running_scripts"
filenm = path + '/' + "monthly_tickets.csv"
outfile = open(filenm, "w")
with outfile:
    writer = csv.writer(outfile, delimiter=',', lineterminator="\n", quoting=csv.QUOTE_NONNUMERIC)
    writer.writerow(col[0] for col in cursor.description)
    for row in cursor:
        writer.writerow(row)

M_Query = cursor.execute(Monthly_tickets).fetchall()

# result=pd.DataFrame(M_Query)
result = pd.DataFrame(M_Query, columns=["INCIDENT_NUMBER", "PRIORITY_CODE", "STATUS", "ASSGNMENT_GROUP", "OPEN_TIME",
                                        "CLOSED_TIME","ASSIGNED_TO", "SPENT_TIME", "SLA"])
result.loc[(result.ASSGNMENT_GROUP == 'UPM (UNT) - OSW'), 'ASSGNMENT_GROUP'] = 'UPM'
result.loc[(result.ASSGNMENT_GROUP == 'MPE - Member Product Eligibility - SRVC'), 'ASSGNMENT_GROUP'] = 'MPE'
result.loc[(result.ASSGNMENT_GROUP == 'Consumer360 - SPT'), 'ASSGNMENT_GROUP'] = 'C360'
result.loc[(result.ASSGNMENT_GROUP == 'BILLING-RECEIVABLE MGMT SYS - APP'), 'ASSGNMENT_GROUP'] = 'AE-MULE'
result.loc[(result.ASSGNMENT_GROUP == 'GATEWAY'), 'ASSGNMENT_GROUP'] = 'FISL'
result.loc[(result.ASSGNMENT_GROUP == 'UHC - OIL API (SUP)'), 'ASSGNMENT_GROUP'] = 'OIL'

# Grouping data
df_res = pd.DataFrame(result, columns=['INCIDENT_NUMBER', 'ASSGNMENT_GROUP', 'SLA'])
# result1
# pd.pivot_table(data=result1,index="ASSGNMENT_GROUP",columns="SLA")
pivot_result = pd.pivot_table(data=df_res, index="ASSGNMENT_GROUP", columns="SLA", fill_value=0,aggfunc=pd.np.count_nonzero, margins=True, margins_name="TOTAL_TICKETS")
#pivot_result=pd.pivot_table(data=df_res,index="ASSGNMENT_GROUP",columns="SLA",fill_value=0,aggfunc=pd.np.count_nonzero,margins=True,margins_name="TOTAL")
pivot_result.columns = pivot_result.columns.droplevel()
pivot_result.columns.name = None

groupwise_COUNT = tabulate(pivot_result.to_records().tolist(), pivot_result.to_records().dtype.names, "html")
groupwise_COUNT

# Plotting

sns.countplot(x='ASSGNMENT_GROUP', hue='SLA', data=result, palette='RdBu_r')
plt.savefig('SLA_Report.png')

# Mailing part

sender = '<pavithramaddi@optum.com>'
receivers = "<pavithramaddi@optum.com>,<pavithramaddi@optum.com>"
####################receivers = "<MPE_Support_DL@ds.uhc.com>,<UPM_Support_DL@ds.uhc.com>"

#month = time.strftime("%B")
month = datetime.strftime(datetime.now() - timedelta(1), '%B')
SUBJECT = "Service-Now Monthly Report - {m}".format(m=month)
message = """
Hello All, </br></br>

<p>Below is the Service-Now report for {m} month.</p> 
The attachment is having the information about the incident ticket with spent time for each ticket and SLA information.</br></br>
The following work-group are included in the report: </br>
AE-MULE, C360, FISL, MPE OIL and UPM </br></br>


    
<table>
<head>
<style>
  table, th, td {{ border: 1px solid black; border-collapse: collapse; padding: 3px;font-family: "verdana";font-size: 12px}}
  th {{background-color: #d4d4aa; line-height: 1.5; text-align: center;}}

</style>
</head>
</table>

{groupwise_COUNT}
</br></br>
<b style="color:#b30047;font-size: 20px;text-align: center;">Group wise SLA Measure</b>
</br>
<img src="SLA_Report.png"></br>
</br></br></br>
<b>Regards,</b></br>
Pavithra</br>
pavithramaddi@optum.com</br></br>

<p>
<b>Note:</b> This is an auto generated mail. Incase of any issues, please drop an email to "pavithramaddi@optum.com"
</p>
""".format(groupwise_COUNT=groupwise_COUNT, m=month)  # ,image_cid=image_cid[1:-1], subtype='html')

"""
with open('savefig.png', 'rb') as img:

    # know the Content-Type of the image
    maintype, subtype = mimetypes.guess_type(img.name)[0].split('/')

    # attach it
    msg.get_payload()[1].add_related(img.read(), 
                                         maintype=maintype, 
                                         subtype=subtype, 
                                         cid=image_cid)"""

############# Mail attachemnt part #######################33

data = MIMEMultipart()
data['From'] = sender
data['To'] = receivers

data['Subject'] = SUBJECT

data.attach(MIMEText(message, 'html'))
# filename = "OPEN-TICKETS"
attachment = open(filenm, "rb")
p = MIMEBase('application', 'octet-stream')
p.set_payload((attachment).read())
encoders.encode_base64(p)
# p.add_header('Content-Disposition', "attachment; filename= %s" % filename)
p.add_header('Content-Disposition', "attachment", filename='monthly_tickets.csv')
readFile = data.attach(p)

# Create our Image Data from the defined image
image = "SLA_Report.png"
imageData = MIMEImage(open(image, 'rb').read(), 'png')
imageData.add_header('Content-Disposition', 'attachment; filename="SLA_Report.png"')
data.attach(imageData)

smtpObj = smtplib.SMTP('mail6.uhc.com')
text = data.as_string()
smtpObj.sendmail(sender, receivers, text, image)
smtpObj.quit()

"""
try:
   smtpObj = smtplib.SMTP('mail6.uhc.com')
   smtpObj.sendmail(sender, receivers, message)
   # smtpObj.sendmail(sender, receivers, message)
   print("Successfully sent email")
except Exception as e:
   print("Error: unable to send email")
   print(e)"""

####################### Closing database #######################
connection.close()



