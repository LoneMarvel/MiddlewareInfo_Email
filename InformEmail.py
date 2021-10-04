from dotenv import load_dotenv
from datetime import datetime
from email.message import EmailMessage
from jinja2 import Template
import smtplib
import os
import gspread
import argparse
import sys

gSheetId = '1_R7n8107iluos2wCSc5WbJRM30LSdOLtlp_xftsKYLg'
baseDir = os.path.abspath(os.path.dirname(__file__))

def OpenWrkSheet():
    gs = gspread.service_account(filename=baseDir+'/creds-files/zcreds.json')    
    gSheet = gs.open_by_key(gSheetId)    
    return gSheet.worksheet('EmailInfo')

def DoReadValues(emailFlag):
    """ 04.10.2021
    Summary
    -----------------------------------
    This function is opening the Google Sheet (MiddlewareReleaseDetails) using OpenWrkSheet function. Reads the Google Sheet and store the data into the following variables
    emailData = Get all the data from Google Sheet
    emailFlag = This is a variable to check if it's the start of the Middleware release or the end
    recipList = This variable is getting the first column from Google Sheet that contains the recipient email list

    Purpose
    -----------------------------------
    Prepare the data for the DoSendEmail function.
    
    Return
    -----------------------------------
    Nothing
    """
    recipList = []
    wrkSheet = OpenWrkSheet()    
    emailData = wrkSheet.get_all_values()
    for i in range(len(emailData)):
        if emailFlag == emailData[i][0]:
            recipList = list(str(emailData[i][1]).split(','))            
            DoSendEmail(recipList,emailData[i][2],emailData[i][3])
            recipList = ''
    return


def DoSendEmail(recipList,subjEmail,msgEmail):
    """ 04.10.2021
    Summary
    -----------------------------------
    This function sends the start/end of Middleware update release. Uses my Zattoo Google Account. The hidden '.env' file contains the credentials data to authenticate and send the email.
    Uses jinja2 template system to open, read and render the html file so the email will be in HTML format. There are two variables that are rendering in the html web page.
    EMAIL_ADDRESS = The content is stored in .env file
    EMAIL_PASSWORD = The content is stored in .env file
    senderName = The content is stored in .env file
    msgEmail = The content is stored in Google Sheet file and is passed from DoReadValues function

    Purpose
    -----------------------------------
    Sends the update/Inform Middleware start/end release update

    Return
    -----------------------------------
    Return nothing
    """
    print(f"Starting")
    load_dotenv(baseDir+'/.env')
    EMAIL_ADDRESS = os.getenv('EMAIL_USER')
    EMAIL_PASSWORD = os.getenv('EMAIL_PASS')
    senderName = os.getenv('SENDER_NAME')
    credsMsg = EmailMessage()      
    openEmailTmpl = open(baseDir+"/templates/IT-EquipmentEmail.html", 'r').read()
    getHtmlCont = Template(openEmailTmpl)  
    msgBody = getHtmlCont.render(msgEmail=msgEmail,senderName=senderName)    
    credsMsg['To'] = recipList
    credsMsg['Subject'] = subjEmail
    credsMsg['From'] = 'noc@zattoo.com'   

    credsMsg.add_alternative(msgBody, subtype='html')            
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        verVar = smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)    
        smtp.send_message(credsMsg)
        print(f"Email To -> {recipList} Sent Successfully . . .")        
    return

my_parser = argparse.ArgumentParser(description='Send Start/End Information Email For Middleware Release Every Tuesday')
my_parser.version = '1.0'
my_parser.add_argument('-s', action='store_true', help='Send Start Information Email For Middleware Release Every Tuesday')
my_parser.add_argument('-e', action='store_true', help='Send End Information Email For Middleware Release Every Tuesday')

if len(sys.argv) < 2:
    print(f'Usage: InformEmail [-s,-e] Or -h For Help')
    exit()

args = my_parser.parse_args()
baseDir = os.path.abspath(os.path.dirname(__file__))

if args.s: DoReadValues('start') 
if args.e: DoReadValues('end')
