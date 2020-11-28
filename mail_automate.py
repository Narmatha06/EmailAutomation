import pandas as pd
import time
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

def send_mails(sender_address,sender_pass,subject,mail_content):
    df = pd.read_excel("data.xlsx")
    for index in df.index:
        receiver_address = df["Email"][index]
        
        #Setup the MIME
        message = MIMEMultipart()
        message['From'] = sender_address
        message['To'] = receiver_address
        
        #The subject line
        message['Subject'] = subject
        
        #The body and the attachments for the mail
        message.attach(MIMEText(mail_content, 'plain'))
        attach_file_name = str(index+1)+'.jpg'
        attach_file = open(attach_file_name, 'rb') # Open the file as binary mode
        payload = MIMEBase('application', 'octate-stream')
        payload.set_payload((attach_file).read())
        
        #encode the attachment
        encoders.encode_base64(payload) 
        
        #add payload header with filename
        payload.add_header('Content-Disposition', "attachment; filename = "+str(df["Name"][index])+".jpg")
        message.attach(payload)
        
        #Create SMTP session for sending the mail
        session = smtplib.SMTP('smtp.gmail.com', 587) #use gmail with port
        session.starttls() #enable security
        session.login(sender_address, sender_pass) #login with mail_id and password
        
        text = message.as_string()
        session.sendmail(sender_address,receiver_address, text)
     
        session.quit()
        print('Mail Sent for '+ df["Name"][index]+" "+str(index+1))
        
if __name__ == "__main__":
    sender_address = 'shabnam91265@gmail.com'
    sender_pass = '9600091265'
    subject = 'Certificate - Webinar on "Trends in Plant Biology"'
    mail_content = '''
                      Warm Greetings from Department of Botany, The American College(Autonomous), Madurai.

                      Dear Participant,

                      Congratulations for completing the Webinar on "Trends in Plant Biology"

                      We have attached the certificate in recognition of your accomplishment. 

                      Note: Correction in the certificate will not be entertained. The details given in the certificate was filled by you in the feedback.

                      Thanks for the support and cooperation.

                      With Regards
                      The Organizing Committee  
                   '''
    send_mails(sender_address,sender_pass,subject,mail_content)
    
