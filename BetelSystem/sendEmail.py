import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import Function


# the function get the email from a txt file, because if you want you sand to some emails...and the anexo is the file
# to send
def send_email(anexo, email):
    # EMAIL AND PASSWORD
    EMAIL_ADRRESS = 'emailaddress@gmail.com'  # Put your email in here
    EMAIL_PASSWORD = 'your_password'  # your email password, you can put a var whot open a file with your password
    EMAIL_RECIVER = email

    # CREATE EMAIL
    msg = MIMEMultipart()
    msg['Subject'] = 'Subject'
    msg['From'] = EMAIL_ADRRESS
    msg['To'] = EMAIL_RECIVER

    # Send
    body = "Message in the body"
    msg.attach(MIMEText(body, 'plain'))
    filename = anexo
    attachment = open(anexo, 'rb')
    p = MIMEBase('application', 'octet-stream')
    p.set_payload((attachment).read())
    encoders.encode_base64(p)

    # fixing the file
    p.add_header('Content-Disposition', "attachment; filename= %s" % filename)
    msg.attach(p)
    s = smtplib.SMTP('smtp.gmail.com', 587)
    s.starttls()
    s.login(EMAIL_ADRRESS, EMAIL_PASSWORD)
    text = msg.as_string()
    s.sendmail(EMAIL_ADRRESS, EMAIL_RECIVER, text)
    s.quit()
