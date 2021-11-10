import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import random
import xlrd
import pandas as pd
import csv

MAIL_SUBJECT = 'Mã OTP gửi từ Tổ Công nghệ - Tin học trường THPT Lê Quý Đôn'

opening_html = '''<p>Chào bạn,
<br>
<br>Đây là email thông báo về Bài kiểm tra Học kì 1 Năm học 2021-2022 môn Tin học.
<br>
<br><strong>Bước 1:</strong> Bạn vào link sau để làm bài
<a href="https://forms.office.com/Pages/ResponsePage.aspx?id=DQSIkWdsW0yxEjajBLZtrQAAAAAAAAAAAAMAAJq0I6FUOEpTSFNUOU1HN0ZUWjg1SzdPUE5KV1dVMS4u" target="_blank">Bắt đầu làm bài</a>
<br>
<br><strong>Bước 2:</strong> Dưới đây là mã OTP dùng để nhập vào bài kiểm tra.
<br>Bạn vui lòng nhập chính xác và không đưa cho người khác.
</p>'''

closing_html = '''<p>Cảm ơn bạn.
<br>Chúc bạn đạt kết quả như ý và thời gian cuối năm thật vui vẻ.
</p>'''

signature_html = '''<span style="color:rgb(61,133,198)"><font size="4">-----------------------------------------------------</font></span>
<figure style="display:flex; align-items: center; justify-content: left;">
<img src="https://lh3.googleusercontent.com/_1Hftqk5e-LXrd3q9KNAsRuH2huEdUEiOPOPzA6IgjnreacL-bL9yAe2Orqc-Dym3NChhogHWHY3_r3UaKVQDBvmhnsRHutWqrc-l98cr4YKOfLbcRbZdvBBae0EmS4yiet37TN2uvg=w120">
    <img src="https://lh3.googleusercontent.com/j9VtcXJ3UlJyOGRCvvvAcZE1seVFj656GTxVyCyu5gml07X-PiHpSLikjKN39Q5pF9_DNk3PoSFHnEGB_E1KFU9FNi9lcaA8BQGXcZYHRXCf2Hr4X5GJeqhzTuowu4DadDnLZgFjgAU=w120">
</figure>
<div><span style="color:rgb(230,145,56)"><b><font size="3">Tổ Công nghệ - Tin học trường THPT Lê Quý Đôn</font></b></span></div>
<div><span style="color:rgb(61,133,198)"><span style="color:rgb(61,133,198)"><font size="2">110 Nguyễn Thị Minh Khai, Phường Võ Thị Sáu, Quận 3</font></span><b><font size="4"><br></font></b></span></div>
'''

mail_content_html = '''\
    <html>
<head></head>
<body>
  <div>{opening_html}
  <h3>{otp_code}</h3>
    {closing_html}
  </div>
   <div>{signature_html}</div> 
  </body>
</html>
'''
#--------------------------------------------------
SENDER = 'some_mail@gmail.com'
PWD = 'some password'
mail_otp_dict = {}

#------------------------------------------
def write_otp_to_csv(otp_dict, csv_dest_file):
    with open(csv_dest_file, 'w') as csvfile:
        csvfile.write("Gmail,OTP\n")
        for key in otp_dict.keys():
            csvfile.write("%s,%s\n" % (key, otp_dict[key]))

#---------------------------------------------
# Create the dict of mail and OTP
def create_mail_otp_dict(receive_mail_csv_file, otp_len):
    #mails = ['vtchitruong@gmail.com', 'krasny019@gmail.com', 'vralph.inc@gmail.com']
    
    df = pd.read_csv(receive_mail_csv_file, usecols = ['Gmail'])
    mails = df['Gmail'].tolist()

    result_dict = {} # Dictionary of mail and OTP
    
    start_number = 10 ** (otp_len - 1)
    end_number = 10 ** otp_len - 1
    
    for m in mails:
        otp = str(random.randint(start_number, end_number))
        result_dict[m] = otp

    # write_otp_to_csv(mail_otp_dict)
    return result_dict

#----------------------------------------------------------------------

# sender email and pass
# sender_mail = 'congdoanthptlequydon.hcm@gmail.com'
# sender_pass = 'zDUvT6@pH7tprfU5M$gW%@qd7'

def start_session(sender, password):
    # create smtp session 
    session = smtplib.SMTP("smtp.gmail.com" , 587)  # 587 is a port number

    # start TLS for E-mail security 
    session.starttls()

    # Log in to your gmail account
    session.login(sender, password)

    return session

#------------------------------------------------------------
def send_to_single_mail(session, sender, mail_subject, recipient, opening, otp_code, closing, signature):
    # setup MIME
    message = MIMEMultipart('Alternative')
    message['From'] = sender
    message['Subject'] = mail_subject
    message['To'] = recipient

    mail_content = mail_content_html.format(opening_html=opening, otp_code = otp_code, closing_html=closing, signature_html=signature)

    message.attach(MIMEText(mail_content, 'html')) # sent in HTML format

    text = message.as_string()       
    session.sendmail(sender, recipient, text)

#------------------------------------------------------------
def close_session(session):
    # close smtp session
    session.quit()

#------------------------------------------------------------
def send_otp(sender, password, recipient_csv_file, mail_subject, otp_length, opening, closing, signature):
    mail_otp_dict = create_mail_otp_dict(recipient_csv_file, otp_length)
    print(mail_otp_dict)
    
    # create smtp session 
    session = smtplib.SMTP("smtp.gmail.com" , 587)  # 587 is a port number

    # start TLS for E-mail security 
    session.starttls()

    # Log in to your gmail account
    session.login(sender, password)

    for r in mail_otp_dict:
        # setup MIME
        message = MIMEMultipart('Alternative')
        message['From'] = sender
        message['Subject'] = mail_subject
        message['To'] = r

        mail_content = mail_content_html.format(opening_html=opening, otp_code = mail_otp_dict[r], closing_html=closing, signature_html=signature)

        message.attach(MIMEText(mail_content, 'html')) # sent in HTML format

        text = message.as_string()       
        session.sendmail(sender, r, text)
        #print(text + '\n')

        #print(content + '\n')
        print("OTP is succesfully sent to " + r)

    # close smtp session
    session.quit()