import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


def enviar_mail(mail, archivo):
    msg = MIMEMultipart()
    msg["Subject"] = "taller.ai"
    msg["From"] = "dontseeeee@gmail.com"
    msg["To"] = mail
    msg.attach(MIMEText(("holaaaaa soy fran (es un mensaje pre escrito perobue), aca esta tu excel o word ns que pediste pero toma bro :D"), "plain"))
    
    with open(archivo, "rb") as f:
        parte = MIMEBase("application", "octet-stream")
        parte.set_payload(f.read())

    encoders.encode_base64(parte)
    parte.add_header(
        "Content-Disposition",
        f"attachment; filename={archivo}"
)

    msg.attach(parte)

    with smtplib.SMTP("smtp.gmail.com", 587) as server:
        server.starttls()
        server.login("dontseee@gmail.com", "shhhhh")
        server.send_message(msg)