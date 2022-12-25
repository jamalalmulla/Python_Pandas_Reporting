import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os.path

reportDate = "27 Jul 2022"
DatePrefix = "20220727_"

#######################################################################################
Teams = ['#Tylus', '#AEREMIAH', '#JasimDelivery','#CLC','#CLC2','#Parcel',"#SWAT","#SWAT_TYLOS","#SWAT_AEREMIAH","#SWAT_JASIM","#SWAT_CLC"]

# ,"#SWAT_TYLOS","#SWAT_AEREMIAH","#SWAT_JASIM","#SWAT_CLC"
for team in Teams:
    to = None
    cc = None
    bcc = None
    attach_file_name = None
    # Please Find attached the daily report.
    # Please Find attached Pation Way Bulk orders report.
    #There is an update in the report. Please consider the new attachement.
    # There was a mistake at the report driver commission at 30 Jun 2022.
    # Please Find attached an updated report for Date 30 Jun 2022.
    mail_content = '''Dear {team},
    
    Please Find attached the daily report.
    
    Thanks and Regards,
    Jamal Al Mulla
    IT Support
    T: +973 - 13333070
    M: +973 - 38388095
    PARCEL 24 DELIVERY CO. W.L.L
    '''.format(team=team)
    #The mail addresses and password
    sender_address = 'jamal@tryparcel.com'
    sender_pass = 'Jamal1984#'

    if team == '#Tylus':
        to = 'hussain.tylos@gmail.com,alseee.tylos@gmail.com'
        cc = 'ali.dhaif@tryparcel.com,hussain.shuaib@tryparcel.com,ammar@tryparcel.com,ali.hussain@tryparcel.com'
        bcc = 'jamal@tryparcel.com'
    elif team == '#AEREMIAH':
        to = 'saeed.aldhaif@gmail.com,Husain.alabbad20@gmail.com'
        cc = 'ali.dhaif@tryparcel.com,hussain.shuaib@tryparcel.com,ammar@tryparcel.com,ali.hussain@tryparcel.com'
        bcc = 'jamal@tryparcel.com'
    elif team == '#JasimDelivery':
        to = 'delivery.aer@gmail.com'
        cc = 'ali.dhaif@tryparcel.com,hussain.shuaib@tryparcel.com,ammar@tryparcel.com,ali.hussain@tryparcel.com'
        bcc = 'jamal@tryparcel.com'
    elif team == '#CLC':
        to = 'ammar@tryparcel.com'
        cc = 'ali.dhaif@tryparcel.com,hussain.shuaib@tryparcel.com,ali.hussain@tryparcel.com,yusuf.alfardan@clc.delivery'
        bcc = 'jamal@tryparcel.com'
    elif team == '#CLC2':
        to = 'ammar@tryparcel.com'
        cc = 'ali.dhaif@tryparcel.com,hussain.shuaib@tryparcel.com,ali.hussain@tryparcel.com,yusuf.alfardan@clc.delivery'
        bcc = 'jamal@tryparcel.com'
    elif team == '#Parcel':
        to = 'ammar@tryparcel.com'
        cc = 'ali.dhaif@tryparcel.com,hussain.shuaib@tryparcel.com,ali.hussain@tryparcel.com,yusuf.alfardan@clc.delivery'
        bcc = 'jamal@tryparcel.com'
    elif team == '#SWAT':
        to = 'ammar@tryparcel.com,jasim@tryparcel.com'
        cc = 'ali.dhaif@tryparcel.com,hussain.shuaib@tryparcel.com,ali.hussain@tryparcel.com'
        bcc = 'jamal@tryparcel.com'
    elif team == '#SWAT_TYLOS':
        to = 'hussain.tylos@gmail.com,alseee.tylos@gmail.com'
        cc = 'ali.dhaif@tryparcel.com,hussain.shuaib@tryparcel.com,ammar@tryparcel.com,ali.hussain@tryparcel.com'
        bcc = 'jamal@tryparcel.com'
    elif team == '#SWAT_AEREMIAH':
        to = 'saeed.aldhaif@gmail.com,Husain.alabbad20@gmail.com'
        cc = 'ali.dhaif@tryparcel.com,hussain.shuaib@tryparcel.com,ammar@tryparcel.com,ali.hussain@tryparcel.com'
        bcc = 'jamal@tryparcel.com'
    elif team == '#SWAT_JASIM':
        to = 'delivery.aer@gmail.com'
        cc = 'ali.dhaif@tryparcel.com,hussain.shuaib@tryparcel.com,ammar@tryparcel.com,ali.hussain@tryparcel.com'
        bcc = 'jamal@tryparcel.com'
    elif team == '#SWAT_CLC':
        to = 'ammar@tryparcel.com'
        cc = 'ali.dhaif@tryparcel.com,hussain.shuaib@tryparcel.com,ali.hussain@tryparcel.com,jasim@tryparcel.com,yusuf.alfardan@clc.delivery'
        bcc = 'jamal@tryparcel.com'



    # to = 'jamal@tryparcel.com'
    # cc = 'jamalalmulla1984@gmail.com'
    # bcc = ''
    # attach_file_name = '{DatePrefix}{team}.xlsx'.format(DatePrefix=DatePrefix, team=team)

# Check if the report exists and attached it.
    if os.path.isfile('{DatePrefix}{team}.xlsx'.format(DatePrefix=DatePrefix, team=team)):
        attach_file_name = '{DatePrefix}{team}.xlsx'.format(DatePrefix=DatePrefix, team=team)
    else:
        print('{team} Report does not exist'.format(team=team))
        to = None
        cc = None
        bcc = None
        attach_file_name = None

    if (to is not None and cc is not None and bcc is not None):
        #Setup the MIME
        message = MIMEMultipart()
        message['From'] = sender_address
        message['To'] = to
        message['Cc'] = cc
        message['Bcc'] = bcc

        rcpt = cc.split(",") + bcc.split(",") + to.split(",")
        # message['Subject'] = team + ' Daily Report ' + reportDate
        # message['Subject'] = team + '- Pation Way Bulk Report ' + reportDate
        message['Subject'] = team + ' Daily Report ' + reportDate
        #The subject line
        #The body and the attachments for the mail
        message.attach(MIMEText(mail_content, 'plain'))

        attach_file = open(attach_file_name, 'rb') # Open the file as binary mode
        payload = MIMEBase('application', 'octate-stream')
        payload.set_payload((attach_file).read())
        encoders.encode_base64(payload) #encode the attachment
        #add payload header with filename
        payload.add_header('Content-Disposition', "attachment; filename= %s" % attach_file_name)
        message.attach(payload)
        #Create SMTP session for sending the mail
        session = smtplib.SMTP('smtp.gmail.com',587) #use gmail with port
        session.starttls() #enable security
        session.login(sender_address, sender_pass) #login with mail_id and password
        text = message.as_string()
        session.sendmail(sender_address, rcpt, text)
        session.quit()
        print(team + ' Report Email Sent Successfully')
    else:
        print(team + ' Report Email Can not be Send. Please check the "to","cc"and"bcc" emails and the report attachement.')
print('The End')