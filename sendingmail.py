'''
Contains a keyword, dropmail, which will senf a mail

Arguments that need to be passed are

tname: Test Case name that failed. 
receivers: receipient email addresses of this email
SMTPhost: SMTP server to be used to send the email
browser: the browser name for which the test case failed. 
env: the environment in which the test case failed. 

Created on 10/10/2012 by Grace
Modifled latest on 11/22/2012 by Grace
'''

from robot.api import logger
from robot.libraries.BuiltIn import BuiltIn
import os
import os.path, time
import sys
import time
import datetime
import imghdr



# Import SMTPLIB for the email sending function
import smtplib
import mimetypes
import base64

# Import the email Module
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.Utils import COMMASPACE, formatdate
from email import Encoders

#from email.message import Message

def dropmail(testCaseName, mailTo, SMTPhost, browser, env, hostname, logFilePath, sourcefile, imagepath):
    
    sender = 'Automation.Tests@ftdi.com'
    mailTo = mailTo.split(',')

    #print 'Receivers from dropmail: [%s]' % ', '.join(map(str, receivers))

    if (testCaseName!=''):                
    	
    	msg = MIMEMultipart() # Define the content type and sender and receiver details
        msg['From']= sender
        msg['To'] = str(COMMASPACE.join(mailTo))
        msg['Subject'] = """FAIL:: %s :: %s :: %s"""%(env,browser,testCaseName)
        msg.attach(MIMEText('Please find attached, files to debug the issue. \nTest was run at : %s'%hostname)) #The content of the email

        #logfile = os.path.join(tempfile.gettempdir(),fname)
        #print ("Log File is at:"+fnamepath)
        
        if (logFilePath!="" and os.path.exists(logFilePath)):
           part = MIMEText(file(logFilePath).read(),"text", 'UTF-8')
           part.add_header('Content-Disposition', 'attachment; filename="%s"' % "Log.txt")
           msg.attach(part) # attaching the log file

        if (sourcefile!="" and os.path.exists(sourcefile)):
           part = MIMEText(file(sourcefile).read(),"html")
           part.add_header('Content-Disposition', 'attachment; filename="%s"' % "HTMLSource.html")        
           msg.attach(part) # attaching the source code file  

        if (imagepath!="" and os.path.exists(imagepath)):
           fp = open(imagepath,'rb')
           img = MIMEImage(fp.read(), name="Screenshot.png")
           fp.close()
           msg.attach(img)           
           
    if (SMTPhost!='' and mailTo!=''):
    	try:
            logger.console("sending mail...")
            smtpObj = smtplib.SMTP(host=SMTPhost, port=25)
            smtpObj.sendmail(sender,mailTo,msg.as_string())  # send the mail
            smtpObj.quit()
            logger.console("...DONE")
            #logger.console("Failure email Sent Successfully")
    	except smtplib.SMTPException as e:
            logger.console("Unable to send email")
            logger.console(str(e))

    	except smtplib.SMTPHeloError as e:
            logger.console("Unable to get reply from sender")
            logger.console(str(e))

    	except smtplib.SMTPRecipientsRefused as e:
            logger.console("Recipients Refused")
            logger.console(str(e))

    	except smtplib.SMTPSenderRefused as e:
            logger.console("Sender Refused")
            logger.console(str(e))

    	except smtplib.SMTPDataError as e:
            logger.console("Data Error")
