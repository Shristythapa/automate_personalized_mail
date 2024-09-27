
from email.mime.text import MIMEText 
from email.mime.image import MIMEImage 
from email.mime.application import MIMEApplication 
from email.mime.multipart import MIMEMultipart 
import smtplib 
import os 
import pandas as pd

df = pd.read_excel("automateemail.xlsx")
df['name'] = df['name'].str.strip()
df['email'] = df['email'].str.strip()
myemail = os.environ.get("email")
mypassword = os.environ.get("password")
print(df.head())
print(myemail)
print(mypassword)
 
def message(name, subject="Python Notification",  
            text="This is python notification", img=None, attachment=None): 
    msg = MIMEMultipart() 
    msg['Subject'] = subject   
    body = f"Dear {name},\n\n{text}"
    msg.attach(MIMEText(body, 'plain'))  

    if img is not None: 
        if type(img) is not list: 
            img = [img]   
        for one_img in img: 
            img_data = open(one_img, 'rb').read()   
            msg.attach(MIMEImage(img_data, name=os.path.basename(one_img))) 

    if attachment is not None: 
        if type(attachment) is not list: 
            attachment = [attachment]   
        for one_attachment in attachment: 
            with open(one_attachment, 'rb') as f: 
                file = MIMEApplication(f.read(), name=os.path.basename(one_attachment)) 
            file['Content-Disposition'] = f'attachment; filename="{os.path.basename(one_attachment)}"' 
            msg.attach(file) 
    return msg 

def mail(): 

    try:
        smtp = smtplib.SMTP_SSL('smtp.gmail.com', 465) 
        smtp.ehlo() 
        smtp.login("****", "****")  
        print("logged")
        for index, row in df.iterrows():
            msg = message(row['name']) 
            print(f"Message created: {msg}")  
            to = row['email']  
            print(f"Sending email to: {to}") 
            
            smtp.sendmail(from_addr="****", to_addrs=to, msg=msg.as_string()) 

    except smtplib.SMTPAuthenticationError:
        print("Error: Authentication failed. Please check your email and password.")
    except smtplib.SMTPException as e:
        print(f"Error sending email: {e}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
    finally:
        smtp.quit()  

if __name__ == "__main__":
    mail()

# do this securely https://sendgrid.com/en-us/blog/smtp-security-and-authentication 