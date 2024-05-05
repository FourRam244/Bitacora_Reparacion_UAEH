import smtplib
import email.mime.multipart
import email.mime.base

#!pip install smtplib

# Crea la conexión SMTP
server = smtplib.SMTP('smtp.gmail.com', 587)

correo = 'uclreparaciones@gmail.com'
pas ='lufs mesw gxyw kvzi'
# Inicia sesión en tu cuenta de Gmail
server.starttls()

server.login(correo, pas)

# Definir el remitente y destinatario del correo electrónico
remitente = "uclreparaciones@gmail.com"
destinatario = "cesaram244@gmail.com"

# Crear el mensaje del correo electrónico
mensaje = email.mime.multipart.MIMEMultipart()
mensaje['From'] = remitente
mensaje['To'] = destinatario
mensaje['Subject'] = "Correo electrónico con archivo adjunto"

# Añadir el cuerpo del mensaje
cuerpo = "Por medio de la presente se le hace llegar el documento comprobante del mantenimiento/reparacion/manufactura de su equipo/pieza, \nCordiales saludos y buen dia \nEn caso de cualquier duda o aclaracion contactarse a: \nreparaciones_ucl@uaeh.edu.mx \n EXT. 13224 "
mensaje.attach(email.mime.text.MIMEText(cuerpo, 'plain'))

# Añadir el archivo como adjunto
ruta_archivo = './1290.pdf'
archivo = open(ruta_archivo, 'rb')
adjunto = email.mime.base.MIMEBase('application', 'octet-stream')
adjunto.set_payload((archivo).read())
email.encoders.encode_base64(adjunto)
adjunto.add_header('Content-Disposition', "attachment; filename= %s" % ruta_archivo)
mensaje.attach(adjunto)

# Convertir el mensaje a texto plano
texto = mensaje.as_string()

# Enviar el correo electrónico
server.sendmail(remitente, destinatario, texto)

# Cerrar la conexión SMTP
server.quit()