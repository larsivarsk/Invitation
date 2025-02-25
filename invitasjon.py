import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import os
from dotenv import load_dotenv
import pandas as pd

load_dotenv()

sender_email = os.environ.get('SENDER_EMAIL')
smtp_server = os.environ.get('SMTP_SERVER')
smtp_port = int(os.environ.get('SMTP_PORT'))
password = os.environ.get('PASSWORD')

test = pd.read_csv('gjesteliste.csv', sep=';')

# Dictionary with full name and email addresses
guest_list = {}
for fornavn, etternavn, mail in zip(test['Fornavn'], test['Etternavn'], test['Mailadresse']):
  full_name = f"{fornavn} {etternavn}"
  name = f"{fornavn}"
  guest_list[full_name] = (mail, name)

with smtplib.SMTP(smtp_server, smtp_port) as server:
  server.starttls()
  server.login(sender_email, password)

  for name, tuple in guest_list.items():
    print(f"Sender til {tuple[1]} på {tuple[0]}")
    email = tuple[0]

    msg = MIMEMultipart("related")
    msg['From'] = sender_email
    msg['To'] = email
    msg['Subject'] = "Du er invitert til fest!"

    # HTML body
    html = f"""<!DOCTYPE html>
    <html>
    <head>
      <title>Email with Embedded Image</title>
      <style>
      body {{
        font-family: serif;
      }}
      </style>  
    </head>
    <body>
      <h3>Kjære {tuple[1]}</h3>
      <table border="0" cellpadding="0" cellspacing="0" width="100%">
      <tr>
        <td align="center">
        <img src="cid:image1" alt="Embedded Image" style="max-width: 100%; height: auto; display: block;">
        </td>
      </tr>
      </table>
      <p>Her er en liste over mulige overnattingsplasser:</p>
      <ul>
      <li><a href="https://www.firsthotels.no/hoteller/norge/lillehammer/first-hotel-breiseth/?gad_source=1&gclid=Cj0KCQiA8fW9BhC8ARIsACwHqYrRfYVOScw-TRpGb7McOKXPYiRBMDk84aBvPp6DY6A5TIuM3VmUoSIaAtpLEALw_wcB">Breiseth Hotell Lillehammer, tlf.: 95979434</a></li>
      <li><a href="https://birkebeineren.no/">Birkebeineren Hotel: tlf.: 61050080 (har også hytteutleie/leiligheter)</a></li>
      <li><a href="https://www.scandichotels.no/hotell/norge/lillehammer/scandic-lillehammer-hotel">Scandic Lillhammer Hotel tlf.: 61286000</a></li>
      <li><a href="https://www.scandichotels.no/hotell/norge/lillehammer/scandic-victoria-lillehammer?utm_source=google&utm_medium=cpc&utm_campaign=NO%20%7C%20Brand%20%7C%20Generic&&cmpid=ppc_BH2d&s_kwcid=AL!7589!3!652699978528!e!!g!!scandic%20victoria%20lillehammer&gad_source=1&gclid=Cj0KCQiA8fW9BhC8ARIsACwHqYpyWfCMu__NtC-YU-Ef_aelLVIH09rqc_bT1fcyVbiSGC-FTBsEofgaAhhKEALw_wcB&gclsrc=aw.ds">Scandic Victoria Hotel tlf.: 61271700</a></li>
      <li><a href="https://www.1881.no">-------og mange flere på 1881.no</a></li>
    </body>
    </html>"""

    msg.attach(MIMEText(html, 'html'))

    with open("Invitasjon.jpg", "rb") as img_file:
      img = MIMEImage(img_file.read())
      img.add_header('Content-ID', '<image1>')  
      img.add_header('Content-Disposition', 'inline', filename='Invitasjon')
      msg.attach(img)

    server.sendmail(sender_email, email, msg.as_string())

print("Alle invitasjoner er sendt!")