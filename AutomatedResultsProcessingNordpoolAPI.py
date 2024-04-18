# title: "AutomatedResultsProcessingNordpoolAPI"
# description: "This script executes the NordpoolAPI console application for results download and export to PIXOS (e.g. 'nordpool_nl')."
# output: ""
# parameters: {}
# owner: "MECO, Lukas Dicke"

"""

Usage:
    AutomatedResultsProcessingNordpoolAPI.py <job_path> --exchangeName=<string>

Options:
    --exchangeName=<string> Matchname of exchange name, see file 'ConfigDataNordpool.xml' in known location.

"""

import subprocess
from subprocess import run
import sys
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

def SendMailPythonServer(send_to, send_cc, send_bcc, subject, body, files=[]):
    msgBody = """<html><head></head>
        <style type = "text/css">
            table, td {height: 3px; font-size: 14px; padding: 5px; border: 1px solid black;}
            td {text-align: left;}
            body {font-size: 12px;font-family:Calibri}
            h2,h3 {font-family:Calibri}
            p {font-size: 14px;font-family:Calibri}
         </style>"""

    msgBody += "<h2>" + subject + "</h2>"
    # msgBody += "<h3>" + message + "</h3>"
    msgBody += body

    strFrom = "no-reply-duswvpyt002p@statkraft.de"
    #strFrom = "nominations@statkraft.de"
    msgRoot = MIMEMultipart('related')
    msgRoot['Subject'] = subject
    msgRoot['From'] = strFrom
    if len(send_to) == 1:
        msgRoot['To'] = send_to[0]
    else:
        msgRoot['To'] = ",".join(send_to)

    if len(send_cc) == 1:
        msgRoot['Cc'] = send_cc[0]
    else:
        msgRoot['Cc'] = ",".join(send_cc)

    if len(send_cc) == 1:
        msgRoot['Bcc'] = send_bcc[0]
    else:
        msgRoot['Bcc'] = ",".join(send_bcc)
    msgRoot.preamble = 'This is a multi-part message in MIME format.'

    msgAlternative = MIMEMultipart('alternative')
    msgRoot.attach(msgAlternative)

    for path in files:
        part = MIMEBase('application', "octet-stream")
        with open(path, 'rb') as file:
            part.set_payload(file.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition',
                        'attachment; filename={}'.format(Path(path).name))
        msgRoot.attach(part)

    msgText = MIMEText('Sorry this mail requires your mail client to allow for HTML e-mails.')
    msgAlternative.attach(msgText)

    msgText = MIMEText(msgBody, 'html')
    msgAlternative.attach(msgText)

    smtp = smtplib.SMTP('smtpdus.energycorp.com')
    smtp.sendmail(strFrom, send_to, msgRoot.as_string())
    smtp.quit()

    print("Mail sent successfully from " + strFrom)


exchangeName = "nordpool_amp4"

exchangeName = str(exchangeName)

recipientsTo = ["lukas.dicke@statkraft.de"]

uncPath = "\\\\energycorp.com\\common\\Divsede\\Operations\\Personal_OPS\\Lukas\\DevelopedApplications\\NordpoolAPI_CWE\\ResultProcessing_NordpoolApi_CWE\\bin\\Debug\\ResultProcessing_NordpoolApi_CWE.exe"

process = subprocess.run(uncPath + " " + exchangeName, capture_output=True)

if process.returncode != 0 :

    error = process.stdout.decode('utf-8') + "<br>Return-code: " + str( process.returncode)

    print(error)

    emailSubject = exchangeName + ": Error when processing result"

    message = exchangeName +": The exchange-result for '" + exchangeName +"' could not be processed, because the following issue occurred:<br><br>" + error
    print(message)


    messageHeader = "Hi colleagues on the intraday-desk,<br><br>please keep track on the issues below. Thanks.<br><br><br>"

else:
    emailSubject = exchangeName + ": Successful result download"
    messageHeader = "Hi colleagues on the intraday-desk,<br><br>"
    message = exchangeName + ": The exchange-result for '" + exchangeName + "' has been successfully processed.<br><br>"
    print("Successful result download!")

messageEnd = "<br><br>BR<br><br>Statkraft Operations"
SendMailPythonServer(send_to=recipientsTo, send_cc=[], send_bcc=[], subject=emailSubject,body=messageHeader + message + messageEnd, files=[])