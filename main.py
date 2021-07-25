import pandas as p
import smtplib as sm
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
data=p.read_excel("email.xlsx")
print (type(data))
email_col=data.get("Emails")
print(email_col)
list_of_emails=list(email_col)
print(list_of_emails)
try:
    server=sm.SMTP("smtp.gmail.com", 587)
    server.starttls()
    server.login("soumyajejyoti1998@gmail.com","Tabun$1998")
    from_="duttasoumyajyoti2015@gmail.com"
    to_=list_of_emails
    message=MIMEMultipart("alternative")
    message['Subject']="Testing my code"
    message["from"]="soumyajejyoti1998@gmail.com"

    html='''
    <html>
    <head>

    </head>
    <body>
        <p>Good night everyone</p>
    </body>
    </html>
    '''

    text=MIMEText(html,"html")
    message.attach(text)
    server.sendmail(from_, to_, message.as_string())
    print("Mission Accomplished")
except Exception as e:
    print(e)