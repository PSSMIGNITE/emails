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
    sender_name = "APMA Info"  # Replace with the name you want to display
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
            <h2>Special Event: Sun Transition Group Meditation</h2>
        </div>
        <div class="container">
            <div class="body-content">
                <div class="text-content" style="margin: 0; line-height: 1.6; background-color: #f5ded5;>
                    <p>Dear {recipient_name},</p>
                    <p>We are thrilled to invite you to our special event filled with meditation, music, and enlightening sessions. Below is the detailed schedule for the day:</p>
                    <p>Here is our upcoming Mega Event Schedule. We come up with some live music and transition sessions ahead of us. </p>
                        <div style="border: 1.5px solid #ffac33; margin-top: 10px;">
            <table style="width: 100%; border-collapse: collapse; margin: auto;">
                <tr>
                    <!-- New Row for Text Content -->
                    <td style="width: 100%; font-family: Arial, sans-serif; text-align: left;">
                        {f'<img class="sun2025" src="cid:sun2025" alt="sun2025" style="width: 100%; height: auto; max-width: 100%;">' if sun2025_path else ''}
                    </td>
                </tr>
                <tr>
                    <!-- New Row for Text Content -->
                    <td style="width: 100%; font-family: Arial, sans-serif; text-align: left;">
                        <div style="margin: 0; line-height: 1.6; background-color: #f5ded5;">
                            <p><strong>Event Details:</strong></p>
                            <p><strong>Event Name:</strong> Sun Transition Group Meditation</p>
                            <p><strong>Date/Time:</strong> Jan 26th/Sunday from <em>11:00 AM to 4:00 PM</em></p>
                            <p><strong>Location:</strong><em>Dream House Studios</em> <br>
                            <a href="https://g.co/kgs/NtDi9Ss"> 56 Perimeter Center E 3D, Atlanta, GA 30346. </a></p>
                            <p><strong>Hosted by:</strong> APMA Team</p>
                            <p> 
                            <a href="https://bit.ly/sun-2025"> <strong>**Register Here**</strong></a>
                            </p>
                            <p>Come and participate in this powerful meditation experience, designed to deepen your practice and contribute energetically to the development of our meditation center.</p> 
                            <p>The session is intended for those who wish to immerse themselves in focused and intentional meditation.</p>
                            <p> 
                                <a href="https://bit.ly/donate-apma"> <strong>**Donate Here**</strong></a>
                            </p>
                        </div>
                    </td>
                </tr>
            </table>
            </div>
            <div style="border: 1.5px; margin-top: 10px;background-color: #f5ded5;">
            <h2>Upcoming Meditation Event Schedule</h2>
                <table class="schedule">
                        <tr>
                            <th>Time</th>
                            <th>Activity</th>
                        </tr>
                        <tr>
                            <td>11:00 – 11:10 AM</td>
                            <td>Welcome Inauguration & Diya Lighting</td>
                        </tr>
                        <tr>
                            <td>11:10 – 11:20 AM</td>
                            <td>Ganesha Song by Tanuja</td>
                        </tr>
                        <tr>
                            <td>11:20 – 11:30 AM</td>
                            <td>Friendship Message by Patriji (Video Presentation)</td>
                        </tr>
                        <tr>
                            <td>11:30 – 11:35 AM</td>
                            <td>Essence of Sun Transition</td>
                        </tr>
                        <tr>
                            <td>11:35 – 12:30 PM</td>
                            <td>Introduction of Jayanth - Flute Performance & Guided Meditation</td>
                        </tr>
                        <tr>
                            <td>12:30 – 1:45 PM</td>
                            <td>Lunch Break</td>
                        </tr>
                        <tr>
                            <td>1:45 – 2:00 PM</td>
                            <td>Dance Performance by Sush & Div</td>
                        </tr>
                        <tr>
                            <td>2:00 – 2:30 PM</td>
                            <td>Breath Awareness by Sirish</td>
                        </tr>
                        <tr>
                            <td>2:30 – 3:00 PM</td>
                            <td>Tips on Parenting by Kristopher Fix</td>
                        </tr>
                        <tr>
                            <td>3:00 – 3:45 PM</td>
                            <td>Sound Bath Meditation</td>
                        </tr>
                        <tr>
                            <td>4:00 PM</td>
                            <td>Announcements & Closing</td>
                        </tr>
                    </table>
            </div>
            <div style="border: 1.5px solid #ffac33; margin-top: 10px;">
            <p>
            <strong>What to Expect:</strong><br>
                - Engage in deep meditation and share your experiences.<br>
                - Enjoy the serene vibrations of pyramid energy.<br>
                - Feel free to bring a yoga mat, cushion, or anything for comfort.
            </p>
            </div>

            <p>If you have any questions or require further details, please feel free to reach out.</p>
            <p>Warm regards,<br>Team APMA<br>Atlanta Pyramid Meditation Academy<br></p>
            </div>
        </div>
     </body>
     
     
    <!-- Footer Section -->
    <footer style="color: #7f8c8d; text-align: center; margin-top: 30px; font-size: 14px; background: linear-gradient(to bottom, #ffffff, #e6e6e6); border-radius: 12px; padding: 0px 20px 20px 20px;">
        <!-- Donation Section -->
        <p>Your contributions help us continue offering free events and programs. If you’d like to support our mission, feel free to donate via <strong>Zeffy</strong> or <strong>Paypal</strong>:</p>
        
        <table width="45%" cellpadding="0" cellspacing="0" border="0" align="center">
            <tr>
                <td class="innerText" style="font-size: 12px; vertical-align: top;">
                    <p><strong>Atlanta Pyramid Meditation Academy</strong><br />Zelle Email: <strong>donate_apma@pssmusa.org</p>
                    <p>&nbsp;</p>
                    <p>
                        <a href="https://www.zeffy.com/en-US/fundraising/donate-to-make-a-difference-4600" target="_blank" style="text-decoration: none; font-weight: bold; color: #007bff;">
                            <img src="https://cdn.brandfetch.io/id_2ESd93s/w/820/h/251/theme/light/logo.png?c=1bx1739594752816id64Mup7ac7THNtNEi&t=1733718796963" alt="Zeffy" width="30" style="vertical-align: middle;"> Donate via Zeffy
                        </a>
                    </p>
                </td>
            </tr>
        </table>
        <p style="font-size: 14px;">Thank you for being part of this incredible journey. Let’s make it even more special together!</p>
    
    
        <!-- Social Media Links -->
        <p>
            <a href="https://instagram.com/atlpma" target="_blank" style="text-decoration: none; margin: 0 5px;">
                <img src="https://upload.wikimedia.org/wikipedia/commons/thumb/6/6e/Instagram_font_awesome.svg/512px-Instagram_font_awesome.svg.png?20170908192522" alt="Instagram" width="32" style="vertical-align: middle;">
            </a>
            <a href="https://facebook.com/atlpma" target="_blank" style="text-decoration: none; margin: 0 5px;">
                <img src="https://upload.wikimedia.org/wikipedia/commons/f/f1/Ec-facebook.png" alt="Facebook" width="30" style="vertical-align: middle;">
            </a>
            <a href="https://x.com/atlpma" target="_blank" style="text-decoration: none; margin: 0 5px;">
                <img src="https://upload.wikimedia.org/wikipedia/commons/thumb/c/ce/X_logo_2023.svg/300px-X_logo_2023.svg.png" alt="Twitter" width="25" style="vertical-align: middle;">
            </a>
            <a href="https://linkedin.com/company/atlpma" target="_blank" style="text-decoration: none; margin: 0 5px;">
                <img src="https://upload.wikimedia.org/wikipedia/commons/thumb/7/76/Font_Awesome_5_brands_linkedin.svg/512px-Font_Awesome_5_brands_linkedin.svg.png?20181017221845" alt="Twitter" width="25" style="vertical-align: middle;">
            </a>
            <a href="https://www.threads.net/@atlpma" target="_blank" style="text-decoration: none;margin: 0 5px;">
                <img src="https://upload.wikimedia.org/wikipedia/commons/d/db/Threads_%28app%29.png" alt="Twitter" width="25" style="vertical-align: middle;">
            </a>
            <a href="https://wa.me/message/2JEYRSBFPLMFN1" target="_blank" style="text-decoration: none; margin: 0 5px;">
                <img src="https://upload.wikimedia.org/wikipedia/commons/thumb/4/4e/Bootstrap_whatsapp.svg/480px-Bootstrap_whatsapp.svg.png" alt="WhatsApp" width="27" style="vertical-align: middle;">
            </a>
            <a href="https://www.youtube.com/@atlpma" target="_blank" style="text-decoration: none;">
                <img src="https://upload.wikimedia.org/wikipedia/commons/thumb/f/fe/YouTube_social_dark_circle_%282017%29.svg/256px-YouTube_social_dark_circle_%282017%29.svg.png?20220808215255" alt="YouTube" width="30" style="vertical-align: middle;">
            </a>
        </p>
        <p>
            For recent events, visit: <a href="http://www.atlpma.org/events" target="_blank" style="text-decoration: none; color: #007bff;">www.atlpma.org/events</a> <br/>
            <a href="https://forms.gle/pegJAt66StFS9Geo6" target="_blank" style="text-decoration: none; color: #007bff;">Subscribe to our weekly newsletter here</a>
        </p>
    </footer>
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