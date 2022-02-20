# Python 3.8.0
from re import S
import smtplib
import time
import imaplib
import email
import traceback 
import xlsxwriter
from datetime import datetime, timedelta
import datetime
# -------------------------------------------------
#
# Utility to read email from Gmail Using Python
#
# ------------------------------------------------
ORG_EMAIL = "@gmail.com" 
FROM_EMAIL = "bvm.msoffice" + ORG_EMAIL 
FROM_PWD = "bvmmemsec@2019" 
SMTP_SERVER = "imap.gmail.com" 
SMTP_PORT = 993

def read_email_from_gmail():
    try:
        mail = imaplib.IMAP4_SSL(SMTP_SERVER)
        mail.login(FROM_EMAIL,FROM_PWD)
        mail.select('inbox')

        data = mail.search(None, 'ALL')
        mail_ids = data[1]
        id_list = mail_ids[0].split()   
        first_email_id = int(id_list[0])
        latest_email_id = int(id_list[-1])
        count = 1
        row = 0
        col = 0
         # import xlsxwriter module

        # Workbook() takes one, non-optional, argument
        # which is the filename that we want to create.
        workbook = xlsxwriter.Workbook('gmaildata.xlsx')

        # The workbook object is then used to add new
        # worksheet via the add_worksheet() method.
        worksheet = workbook.add_worksheet()
        for i in range(latest_email_id,latest_email_id-2, -1):
            data = mail.fetch(str(i), '(RFC822)' )
            for response_part in data:
                arr = response_part[0]
                if isinstance(arr, tuple):
                    msg = email.message_from_string(str(arr[1],'utf-8'))
                
                    email_subject = msg['subject']
                    email_from = msg['from']
                    start_email = email_from.find("<")
                    end_email = email_from.find(">")
                    fr_email = email_from[start_email+1:end_email]
                    dt = msg['Date']
                    dt_index = dt.find(':')
                    dt = dt[0:dt_index-3]
                
                    print(row, '--------------------------')
                    print('Date:', dt)
                    print('From : ' + fr_email + '\n')
                    print('Subject : ' + email_subject + '\n')
                  
                    worksheet.write(row, col, row+1)
                    worksheet.write(row, col + 1, dt)
                    worksheet.write(row, col + 2, fr_email)
                    worksheet.write(row, col + 3, email_subject)
                    # worksheet.write('D1', 'Geeks')
                    row +=1
                    # count = count + 1
                    # Finally, close the Excel file
                    # via the close() method.
            if i >1:
                break
        workbook.close()

    except Exception as e:
        traceback.print_exc() 
        print(str(e))

read_email_from_gmail()