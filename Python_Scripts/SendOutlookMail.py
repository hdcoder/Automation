import win32com.client
import csv

def send_mail_via_com(text, subject, recipient, profilename="Outlook2016"):
    o = win32com.client.Dispatch("outlook.application")

    Msg = o.CreateItem(0)
    Msg.To = recipient

    # Msg.CC = "moreaddresses here"
    # Msg.BCC = "address"

    Msg.Subject = subject
    Msg.HTMLBody = text

    # attachment1 = "Path to attachment no. 1"
    # attachment2 = "Path to attachment no. 2"
    # Msg.Attachments.Add(attachment1)
    # Msg.Attachments.Add(attachment2)

    Msg.Send()


def sendMail(name,gender,email):
    send_mail_via_com(getMailBody(name,gender),"Fixed subject",email)

def getMailBody(name,gender):
    mailBody = """\
<html>
  <head></head>
  <body>
    <h3>Hi,
    """ + getSalutation(gender) + " " + name + """</h3>\<br>
       <p>
       How are you?<br>
       Here is the <a href="http://www.python.org">link</a> you wanted.
       </p>
  </body>
</html>
"""
    return mailBody


def getSalutation(gender):
    if(gender=="Male"):
        return "Mr."
    elif(gender=="Female"):
        return "Ms."


f = open('dummy_names.csv')
csv_f = csv.reader(f)

for row in csv_f:
  print row[0]
  sendMail(row[0],row[1],row[2])