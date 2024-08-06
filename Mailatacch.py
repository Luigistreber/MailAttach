#!/usr/bin/env python3
import imaplib
import email
from email.policy import default
import os
import re
from datetime import datetime
import pandas as pd
import PyPDF2
from docx import Document

# Configuración del servidor de correo y las credenciales
IMAP_SERVER = 'imap.gmail.com'
EMAIL_ACCOUNT = 'lbstreber@gmail.com'
PASSWORD = 'ncxc eajm nkzm qiph'  # Usa la contraseña de aplicación generada aquí

# Carpeta donde guardar las imágenes
SAVE_FOLDER = 'imagenes_correo'

# Archivo Excel con la base de datos
EXCEL_FILE = 'pruebapy.xlsx'

# Crear la carpeta si no existe
if not os.path.exists(SAVE_FOLDER):
    os.makedirs(SAVE_FOLDER)

def connect_to_email():
    mail = imaplib.IMAP4_SSL(IMAP_SERVER)
    mail.login(EMAIL_ACCOUNT, PASSWORD)
    mail.select('inbox')  # Conectar a la bandeja de entrada
    return mail

def search_emails_between_dates(mail, start_date, end_date):
    start_date_str = start_date.strftime('%d-%b-%Y')
    end_date_str = end_date.strftime('%d-%b-%Y')
    result, data = mail.search(None, f'SINCE {start_date_str} BEFORE {end_date_str}')
    if result == 'OK':
        return data[0].split()
    else:
        return []


def get_name_from_email(email_address, excel_data):
    # Convertir la dirección de correo electrónico a minúsculas
    email_address_lower = email_address.lower()
    # Buscar el correo electrónico en la columna 'email' del DataFrame
    # y obtener el nombre correspondiente en la misma fila
    row = excel_data[excel_data['email'].str.lower() == email_address_lower]
    if not row.empty:
        return row['name'].iloc[0]  # Obtener el nombre de la primera fila coincidente
    return None

def download_attachments(mail, email_ids, excel_data):
    for email_id in email_ids:
        result, data = mail.fetch(email_id, '(RFC822)')
        if result == 'OK':
            email_message = email.message_from_bytes(data[0][1], policy=default)
            sender_email = email.utils.parseaddr(email_message['From'])[1]
            sender_name = get_name_from_email(sender_email, excel_data)

            if sender_name:
                for part in email_message.walk():
                    if part.get_content_maintype() == 'image':
                        # Descargar imágenes adjuntas
                        filename = part.get_filename()
                        if filename:
                            file_extension = os.path.splitext(filename)[1]
                            new_filename = f"{sender_name}{file_extension}"
                            filepath = os.path.join(SAVE_FOLDER, new_filename)
                            with open(filepath, 'wb') as f:
                                f.write(part.get_payload(decode=True))
                            print(f'Imagen guardada: {filepath}')
                    elif part.get_content_maintype() == 'application':
                        # Descargar y procesar archivos PDF y documentos Word
                        if part.get_content_subtype() == 'pdf':
                            filename = part.get_filename()
                            if filename:
                                file_extension = os.path.splitext(filename)[1]
                                new_filename = f"{sender_name}{file_extension}"
                                filepath = os.path.join(SAVE_FOLDER, new_filename)
                                with open(filepath, 'wb') as f:
                                    f.write(part.get_payload(decode=True))
                                print(f'PDF guardado: {filepath}')
                                extract_images_from_pdf(filepath, sender_name)
                        elif part.get_content_subtype() in ['msword', 'vnd.openxmlformats-officedocument.wordprocessingml.document']:
                            filename = part.get_filename()
                            if filename:
                                file_extension = os.path.splitext(filename)[1]
                                new_filename = f"{sender_name}{file_extension}"
                                filepath = os.path.join(SAVE_FOLDER, new_filename)
                                with open(filepath, 'wb') as f:
                                    f.write(part.get_payload(decode=True))
                                print(f'Documento Word guardado: {filepath}')
                                extract_images_from_docx(filepath, sender_name)
            else:
                print(f'No se encontró el nombre para el correo: {sender_email}')

def extract_images_from_pdf(pdf_filepath, sender_name):
    with open(pdf_filepath, 'rb') as f:
        reader = PyPDF2.PdfFileReader(f)
        for page_num in range(reader.numPages):
            page = reader.getPage(page_num)
            if '/XObject' in page['/Resources']:
                xObject = page['/Resources']['/XObject'].getObject()
                for obj in xObject:
                    if xObject[obj]['/Subtype'] == '/Image':
                        size = (xObject[obj]['/Width'], xObject[obj]['/Height'])
                        data = xObject[obj]._data
                        if xObject[obj]['/ColorSpace'] == '/DeviceRGB':
                            mode = "RGB"
                        else:
                            mode = "P"
                        if '/Filter' in xObject[obj]:
                            if xObject[obj]['/Filter'] == '/FlateDecode':
                                img = Image.frombytes(mode, size, data)
                                img_filename = f"{sender_name}_page{page_num}_image{obj[1:]}.png"
                                img.save(os.path.join(SAVE_FOLDER, img_filename))
                                print(f'Imagen extraída del PDF y guardada: {img_filename}')

def extract_images_from_docx(docx_filepath, sender_name):
    document = Document(docx_filepath)
    for idx, image in enumerate(document.inline_shapes):
        # Extraer la imagen del documento Word
        image_bytes = image._inline.graphic.graphicData.pic.blipFill.blip.embed._blob
        image_filename = f"{sender_name}_image{idx}.png"
        image_filepath = os.path.join(SAVE_FOLDER, image_filename)
        with open(image_filepath, 'wb') as f:
            f.write(image_bytes)
            print(f'Imagen extraída del documento Word y guardada: {image_filename}')

def main():
    # Conectar al servidor de correo
    mail = connect_to_email()
    
    # Leer la base de datos Excel
    excel_data = pd.read_excel(EXCEL_FILE)
    
    # Solicitar fechas de inicio y finalización del usuario
    start_date_str = input("Introduce la fecha de inicio (formato YYYY-MM-DD): ")
    end_date_str = input("Introduce la fecha de finalización (formato YYYY-MM-DD): ")
    
    start_date = datetime.strptime(start_date_str, '%Y-%m-%d')
    end_date = datetime.strptime(end_date_str, '%Y-%m-%d')
    
    # Buscar correos entre las fechas especificadas
    email_ids = search_emails_between_dates(mail, start_date, end_date)
    
    # Descargar adjuntos (imágenes) y renombrarlos según el nombre del remitente
    download_attachments(mail, email_ids, excel_data)
    
    # Cerrar la conexión
    mail.logout()

if __name__ == '__main__':
    main()


