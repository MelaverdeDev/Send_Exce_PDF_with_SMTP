import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from win32com import client
import os

#Settings for file excel and PDF
WB_PATH = "excel" #Path to original excel file
PATH_TO_PDF = "yourpdf.pdf" #PDF path when saving
path = os.path.abspath(WB_PATH)
pathDest = os.path.abspath(PATH_TO_PDF)

#Setting for SMTP
sender_email = "sender_email"
sender_password = "sender_password"
receiver_email = "receiver_email"
subject = "Subject"
body = "Body mail message"
attachment_path = pathDest
attachment_name = "filename.pdf"
smtp_server = "smtp.server"
smtp_port = 25



def send_email_exch(sender_email, sender_password, receiver_email, subject, body, attachment_path, attachment_name, smtp_server, smtp_port):
    #Create an object MIME
    msg = MIMEMultipart()
    msg['Subject'] = subject
    msg['From'] = sender_email
    msg['To'] = receiver_email
    
    #Add text message to body
    text_part = MIMEText(body, 'plain')
    msg.attach(text_part)
    
    with open(attachment_path, 'rb') as attachment:
        attachment_part = MIMEApplication(attachment.read())
        attachment_part.add_header('Content-Disposition', 'attachment', filename=attachment_name)
        msg.attach(attachment_part)

    
    try:
        # Create a connecction SMTP
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.ehlo()
        print(server.help())

        # Debug option
        server.set_debuglevel(False)

        # Auth
        server.login(sender_email, sender_password)

        # Send Email
        server.sendmail(sender_email, receiver_email, msg.as_string())

        # Close connection
        server.quit()

        print("Email inviata correttamente")
    except Exception as e:
        print("Errore durante l'invio dell'email:", str(e))

#Excel part
excel = client.DispatchEx("Excel.Application")
excel.Interactive = False
excel.Visible = False
workbook = excel.Workbooks.Open(path)
 
# Refresh all data connections if you have external connections
workbook.RefreshAll()
excel.CalculateUntilAsyncQueriesDone()
workbook.Save()
print("Converting PDF")
workbook.ActiveSheet.ExportAsFixedFormat(0,pathDest)
workbook.Close()
excel.Quit()

print("Completed")



#Send email using an SNMP server
send_email_exch(sender_email, sender_password, receiver_email, subject, body, attachment_path, smtp_server, smtp_port)


