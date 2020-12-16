# Importar las librerias y paquetes necesarios
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import smtplib
from variables import COTIZACIONES_PATH


def send_cotizacion(cotizacion):
    # Crear el objeto del mensaje
    msg = MIMEMultipart()

    # Escribir el cuerpo del mensaje
    body = "Apreciado cliente, adjunto envio respuesta de su cotización, si tiene alguna inquietud, " \
           "no dude en contactarse con nosotros" \
           "\n" \
           "\n" \
           "QuailHouse"

    # Configurar los parámetros adicionales del objeto y correo
    password = "compu20graph20"
    msg['From'] = "compugraficaresponse@gmail.com"
    msg['To'] = cotizacion["correo"]
    msg['Subject'] = "Respuesta Solicitud de Cotización"

    if cotizacion["es_invernadero"]:
        file_paths = [COTIZACIONES_PATH + "\\modelo3d.pdf"]
    else:
        file_paths = [COTIZACIONES_PATH + "\\Alimentadora_de_codornices.pdf"]

    for path in file_paths:
        # Buscar ruta de los archivos a adjuntar y seleccionar cuales archivos
        # file_name = 'cotizacion.pdf'
        file_name = path.split("\\")[-1]

        # Agregar el cuerpo al mensaje
        msg.attach(MIMEText(body, 'plain'))

        # Abrimos el archivo que vamos a adjuntar
        archivo_adjunto = open(path, 'rb')

        # Creamos un objeto MIME base
        adjunto_MIME = MIMEBase('application', 'octet-stream')
        # Y le cargamos el archivo adjunto
        adjunto_MIME.set_payload((archivo_adjunto).read())
        # Codificamos el objeto en BASE64
        encoders.encode_base64(adjunto_MIME)
        # Agregamos una cabecera al objeto
        adjunto_MIME.add_header('Content-Disposition', "attachment; filename= %s" % file_name)
        # Y finalmente lo agregamos al mensaje
        msg.attach(adjunto_MIME)

    # Creamos el servidor del correo
    server = smtplib.SMTP('smtp.gmail.com: 587')

    server.starttls()

    # Ingresamos al correo electrónico con las credenciales
    server.login(msg['From'], password)

    # Envio del mensaje
    server.sendmail(msg['From'], msg['To'], msg.as_string())

    server.quit()

    print("Correo enviado exitosamente")
