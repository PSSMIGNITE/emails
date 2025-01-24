import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.utils import formataddr
import os
import pandas as pd

def send_email(subject, body, to_email, from_email, from_password, recipient_name, flyer1_path=None, file_path=None):
    # Create message container with HTML
    msg = MIMEMultipart()
    # Set the sender's name and email in the "From" field
    sender_name = "TEAM PSSM"  # Replace with the name you want to display
    msg['From'] = formataddr((sender_name, from_email))
    msg['To'] = to_email
    msg['Subject'] = subject

    # Define HTML content for the email with an embedded logo

    weekly_html_content = f"""
    <html>
        <head>
    <title>Event Schedule</title>
    <style>
        body {{
            font-family: Arial, sans-serif;
            line-height: 1.6;
            color: #333;
        }}
        h1, h2 {{
            color: #4CAF50;
            font-size: 1.4em;
            text-align: center;
        }}
        .schedule {{
            border-collapse: collapse;
            width: 100%;
            margin-top: 20px;
        }}
        .schedule th, .schedule td {{
            border: 1px solid #ddd;
            padding: 8px;
            text-align: left;
        }}
        .schedule th {{
            background-color: #f4f4f4;
        }}
        .button {{
            background-color: #4CAF50;
            color: white;
            padding: 10px 20px;
            text-decoration: none;
            border-radius: 5px;
            display: inline-block;
        }}
        .button:hover   {{
            background-color: #45a049;
        }}
        .footer {{
            margin-top: 20px;
            font-size: 0.9em;
            color: #666;
        }}
    </style>
</head>
     <body>
        <div class="header">
            {f'<img class = "banner" src="cid:flyer1" alt="Organization Logo" width="100%" max-width="100%">' if flyer1_path else ''}
            <h1>Join Us for a Day of Meditation, Music, and Mindfulness</h1>
            <h2>Become a Master - IN PERSON WORKSHOP </h2>
        </div>
        <div class="container">
            <div class="body-content">
                <div class="text-content" style="margin: 0; line-height: 1.6; background-color: #f5ded5;>
                    <p>Dear {recipient_name},</p>
                    <p>We are thrilled to invite you to our special event filled with meditation, music, and enlightening sessions. Below is the detailed schedule for the day:</p>
                    <p>Here is our upcoming Mega Event Schedule. We come up with some live music and transition sessions ahead of us. </p>
                </div>
            </div>
            <table style="width: 100%; border-collapse: collapse; margin: auto;">
               
                <tr>
                    <!-- New Row for Text Content -->
                    <td style="width: 100%; font-family: Arial, sans-serif; text-align: left;">
                        <div style="margin: 0; line-height: 1.6; background-color: #f5ded5;">
                            <p><strong>Event Details:</strong></p>
                            <p><strong>Event Name:</strong> Becoming a Master - InPerson Workshop</p>
                            <p><strong>Date/Time:</strong> Feb 22nd/Saturday from <em>09:00 AM to 05:00 PM</em></p>
                            <p><strong>Location:</strong><em>DreamDestination Ranch </em> <br>
                            <a href="https://maps.app.goo.gl/dDa3E2s99kVzx1H4A"> 3022 Rockhill Rd, Aubrey, TX 76227. </a></p>
                            <p><strong>Hosted by:</strong> Team PSSM </p>
                            <p> 
                            <a href="https://bit.ly/PSSM18principles"> <strong>**Register Here**</strong></a>
                            </p>
                            <p>Come and participate in this powerful meditation experience, designed to deepen your practice and contribute energetically to the development of our meditation center.</p> 
                            <p>The session is intended for those who wish to immerse themselves in focused and intentional meditation.</p>
                        
                        </div>
                    </td>
                </tr>
                 <tr>
                    <!-- New Row for Text Content -->
                    <td style="width: 100%; font-family: Arial, sans-serif; text-align: left;">
                        {f'<img class="sun2025" src="cid:sun2025" alt="sun2025" style="width: 100%; height: auto; max-width: 100%;">' if sun2025_path else ''}
                    </td>
                </tr>
            </table>
           
           
            <div style="border: 1.5px solid #ffac33; margin-top: 10px;">
            <p>
            <strong>What to Expect:</strong><br>
                - Engage in deep meditation and share your experiences.<br>
                - Enjoy the serene vibrations of pyramid energy.<br>
                - Feel free to bring a yoga mat, cushion, or anything for comfort.
            </p>
            </div>

            <p>If you have any questions or require further details, please feel free to reply to this email.</p>
            <p>Contact - 490 793 8380 , 631 974 4518 , 901 212 5834  .</p>
            <p>Warm regards,<br>Team PSSM <br>With Divine Guidance from Patriji <br></p>
            </div>
        </div>
     </body>
     
     
   
    </html>
    """

    # Attach the HTML content to the message
    msg.attach(MIMEText(weekly_html_content, "html"))

    # Attach the logo as an inline image if provided
    if flyer1_path and os.path.isfile(flyer1_path):
        with open(flyer1_path, 'rb') as flyer1:
            flyer1_part = MIMEBase('application', 'octet-stream')
            flyer1_part.set_payload(flyer1.read())
            encoders.encode_base64(flyer1_part)
            flyer1_part.add_header('Content-ID', '<flyer1>')
            flyer1_part.add_header('Content-Disposition', 'inline', filename=os.path.basename(flyer1_path))
            msg.attach(flyer1_part)

    # Attach the flyer5 as an inline image if provided
    if sun2025_path and os.path.isfile(sun2025_path):
        with open(sun2025_path, 'rb') as sun2025:
            sun2025_part = MIMEBase('application', 'octet-stream')
            sun2025_part.set_payload(sun2025.read())
            encoders.encode_base64(sun2025_part)
            sun2025_part.add_header('Content-ID', '<sun2025>')
            sun2025_part.add_header('Content-Disposition', 'inline', filename=os.path.basename(sun2025_path))
            msg.attach(sun2025_part)

    server = None  # Initialize server as None to avoid UnboundLocalError
    try:
        # Set up the SMTP server using SSL
        server = smtplib.SMTP_SSL("mail.privateemail.com",465)  # Use your provided SMTP server and port
        server.login(from_email, from_password)

        # Send the email
        server.sendmail(from_email, to_email, msg.as_string())
        print("Email sent successfully to !", to_email)

    except Exception as e:
        print(f"Failed to send email: {e}")

    finally:
        # Quit server if it was successfully created
        if server:
            server.quit()

# Read contacts from Excel file
csv_file = "attachments/Test.csv"
#csv_file = "attachments/contacts.csv"

try:
    contacts = pd.read_csv(csv_file)

    # Usage example
    from_email = "registrations@pssmignite.org"
    from_password = "Patriji"
    if os.name == "posix":
        flyer1_path = "images/FEB2025/Flyer01.png"
        sun2025_path = "images/FEB2025/MasterFlyer01.png"

    # Loop through each contact and send an email
    for index, row in contacts.iterrows():
        name = row['FirstName']
        email = row['Email']
        subject = "Becoming a Master - Inperson Telugu Workshop "
        send_email(subject, "", email, from_email, from_password, name, flyer1_path, "")
except Exception as e:
    print(f"Error reading Excel file: {e}")