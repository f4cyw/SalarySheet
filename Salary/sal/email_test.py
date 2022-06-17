import os
import smtplib
 
from email import encoders
from email.utils import formataddr
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
 
from_addr = formataddr(('오정훈', 'bst524@bigstonestudio.com'))
to_addr = formataddr(('오정훈', 'f4cyw@naver.com'))
 
session = None
try:
    # SMTP 세션 생성
    session = smtplib.SMTP('smtp.gmail.com', 587)
    session.set_debuglevel(True)
    
    # SMTP 계정 인증 설정
    session.ehlo()
    session.starttls()
    session.login('j.h.empions@gmail.com', 'wyinguhzfbxulzvi')
 
    # 메일 콘텐츠 설정∏
    message = MIMEMultipart("mixed")
    
    # 메일 송/수신 옵션 설정
    message.set_charset('utf-8')
    message['From'] = from_addr
    message['To'] = to_addr
    message['Subject'] = '안녕하세요'
 
    # 메일 콘텐츠 - 내용
    body = '''
    <h2>안녕하세요.</h1>
    <h4>한달동안 수고 많으셨습니다.</h1>
    <h4>급여 명세서 확인하십시오.</h1>
    <h4>오정훈 배상.</h1>
    '''
    bodyPart = MIMEText(body, 'html', 'utf-8')
    message.attach( bodyPart )
 
    # 메일 콘텐츠 - 첨부파일
    attachments = [
        # os.path.join( os.getcwd(), 'storage', '/Users/jonathanoh/Desktop/python/Salary/xltopdf/고지석_12월급여.pdf' )
        os.path.join('/Users/jonathanoh/Desktop/python/Salary/xltopdf/고지석_12월급여.pdf' )
    ]
 
    for attachment in attachments:
        attach_binary = MIMEBase("application", "octect-stream")
        try:
            binary = open(attachment, "rb").read() # read file to bytes
 
            attach_binary.set_payload( binary )
            encoders.encode_base64( attach_binary ) # Content-Transfer-Encoding: base64
            
            filename = os.path.basename( attachment )
            attach_binary.add_header("Content-Disposition", 'attachment', filename=('utf-8', '', filename))
            
            message.attach( attach_binary )
        except Exception as e:
            print( e )
 
    # 메일 발송
    session.sendmail(from_addr, to_addr, message.as_string())        
 
    print( 'Successfully sent the mail!!!' )
except Exception as e:
    print( e )
finally:
    if session is not None:
        session.quit()
