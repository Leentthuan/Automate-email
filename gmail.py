import pandas as pd
import smtplib
import ssl
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.message import EmailMessage

# Attach message body content
# Create the body of the message (a plain-text and an HTML version).
df = pd.read_excel('/Users/user1/Desktop/Sheet2.xlsx')
df['Duration'] = df['Duration'].dt.strftime('%B %d, %Y')

for r in df.itertuples():      
        First_name = r[1]
        Name = r[2]
        Recipient = r[3]
        Description = r[4]
        Discount_code = r[5]
        Discount_amount = r[6]
        Duration = r[7]
        Email_sent = r[8]
        
# Can change HTML message depending on the sender
        msg = f"""
        Dear {r[1]},
        <br>Here is  your code {r[5]} for {r[4]}.
        <br>Use this to get {r[6]} discount on your next purchase.
        <br>The code will expire after {r[7]}. 
        <br>Best regard,
        <br>Lee.

        """
        print(msg)
# Define the transport variables
        ctx = ssl.create_default_context()
        password = "your app password"    # Your app password goes here
# Generate app password:"https://support.google.com/mail/answer/185833?hl=en-GB"
        sender = "Youremail@gmail.com"   # Your e-mail address
        receiver = r[3] # Recipient's address
        print(receiver)
        
# Create the message
        message = MIMEMultipart("mixed")
        message["From"] = sender
        message["To"] = receiver
        message["Subject"] = "!Discount?!"

# Record the MIME types of both parts - text/plain and text/html.
        #part1 = MIMEText(text, 'plain')
        part1 = MIMEText(msg, 'html')

# Attach parts into message container.
        # According to RFC 2046, the last part of a multipart message, in this case
        # the HTML message, is best and preferred.
        #message.attach(part1)
        message.attach(part1)

# Attach image
        ##filename = '/Users/user1/Desktop/Sheet2.xlsx'
        ##with open(filename, "rb") as f:
        ##    file = MIMEApplication(f.read())
                
        ##disposition = f"attachment; filename={filename}"
        ##file.add_header("Content-Disposition", disposition)
        ##message.attach(file)

# Connect with server and send the message
        with smtplib.SMTP_SSL("smtp.gmail.com", port=465, context=ctx) as server:
                server.login(sender, password)
                server.sendmail(sender, receiver, message.as_string())
                server.quit()
                print("Mail sent!")
