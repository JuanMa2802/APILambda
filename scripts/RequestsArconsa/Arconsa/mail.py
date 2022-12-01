from office365.runtime.auth.client_credential import ClientCredential
from office365.runtime.http.request_options import RequestOptions
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
from office365.sharepoint.listitems.listitem import ListItem
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import time
import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

def mail(email):
    session = smtplib.SMTP("smtp.office365.com:587")
    session.starttls() #Puts connection to SMTP server in TLS mode
    session.login('automations@arconsa.com.co', 'Lambda63074')
    Subject=f'Reporte de orden de compra '+datetime.datetime.now().strftime('%Y-%m-%d %H h %M m')
    remitente = 'automations@arconsa.com.co'
    msgRoot = MIMEMultipart('related')
    msgRoot['Subject'] =  Subject
    msgRoot['From'] = remitente
    msgRoot['To'] =email
    msgRoot.preamble = 'This is a multi-part message in MIME format.'
    msgAlternative = MIMEMultipart('alternative') 
    msgRoot.attach(msgAlternative) 
    msgText ='<h3>A continuación se adjunta el reporte de la orden de compra: </h3><br>'
    msgAlternative.attach(MIMEText(msgText, 'html'))
    nombre_adjunto = 'Reporte de orden de compra '+datetime.datetime.now().strftime('%Y-%m-%d %H h %M m')+'.xlsx'
    time.sleep(1)
    archivo_adjunto = open(r"C:\Lambda Analytics\RequestsArconsa\Reportes\pedidosOrden.xlsx", 'rb')
    time.sleep(3)
    adjunto_MIME = MIMEBase('application', 'octet-stream')
    adjunto_MIME.set_payload((archivo_adjunto).read())
    encoders.encode_base64(adjunto_MIME)
    adjunto_MIME.add_header('Content-Disposition', "attachment; filename= %s" % nombre_adjunto)
    msgRoot.attach(adjunto_MIME)
    session.sendmail(msgRoot['From'], msgRoot['To'], msgRoot.as_string())
    session.quit()
