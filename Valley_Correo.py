import csv
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Detalles del correo
correo_saliente = 'jordan@valleyintlsupply.com'
contraseña = '##'
smtp_host = 'smtp.office365.com'
smtp_port = 587

# Leer csv con los datos del cliente
with open(r'C:\Users\jorda\Downloads\Clientes.csv', newline='') as archivo_csv:
    lector_csv = csv.DictReader(archivo_csv)
    for fila in lector_csv:
        destinatario = fila['Correo']
        idioma = fila['Idioma']

        mensaje = MIMEMultipart()
        mensaje['From'] = correo_saliente

        if idioma == 'Espanol':
            mensaje['Subject'] = 'Pendiente'
            mensaje_correo = 'Pendiente.'
        elif idioma == 'Ingles':
            mensaje['Subject'] = 'Pendiente'
            mensaje_correo = 'Pendiente'
        else:
            continue

        mensaje['To'] = destinatario
        mensaje.attach(MIMEText(mensaje_correo, 'plain'))

        # Iniciar sesión en el servidor SMTP
        servidor_smtp = smtplib.SMTP(smtp_host, smtp_port)
        servidor_smtp.starttls()

        # Autenticar con la cuenta de correo saliente
        servidor_smtp.login(correo_saliente, contraseña)

        # Enviar el correo electrónico
        servidor_smtp.send_message(mensaje)

        # Cerrar la conexión con el servidor SMTP
        servidor_smtp.quit()

        print(f"Correo enviado a: {destinatario}")