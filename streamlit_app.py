import streamlit as st
import pandas as pd
import re
from datetime import datetime, timedelta
import pytz
import win32com.client
import os
import PyPDF2
from openpyxl import load_workbook

st.title("🎈 My new app")
st.write(
    "Let's start building! For help and inspiration, head over to [docs.streamlit.io](https://docs.streamlit.io/)."
)


# Obtener el directorio de documentos del usuario
documents_dir = os.path.join(os.path.expanduser("~"), "Documents")

# Asegurarse de que el directorio de documentos exista
if not os.path.exists(documents_dir):
    os.makedirs(documents_dir)

# Solicitar al usuario la carpeta dentro de "Bandeja de entrada"
carpeta_correos = input("Ingrese el nombre de la carpeta dentro de 'Bandeja de entrada' para buscar correos: ")

# Solicitar al usuario el número de días hacia atrás para revisar los correos
dias_atras = int(input("Ingrese el número de días hacia atrás para revisar los correos: "))

# Solicitar al usuario la carpeta de descarga de archivos PDF
carpeta_descarga = input("Ingrese la carpeta donde desea guardar los archivos PDF (relativo a Documentos): ")
actas_pdf_dir = os.path.join(documents_dir, carpeta_descarga)

# Crear la carpeta si no existe
if not os.path.exists(actas_pdf_dir):
    os.makedirs(actas_pdf_dir)
    print(f"La carpeta se ha creado: {actas_pdf_dir}")

# Solicitar al usuario el correo al que se debe enviar el archivo procesado
correo_destino = input("Ingrese el correo electrónico de destino: ")

# Configuración
subject_keyword = "Certificado de entrega  ODV-"
save_dir = actas_pdf_dir
fecha_actual = datetime.now().date()

# Formatear la fecha con el formato deseado
fecha_formateada = fecha_actual.strftime("%d-%m-%y")

# Lista de prompts
PROMPT_LIST = [
    {'BATCH FECHA INICIO': r'BATCH FECHA INICIO\s*(\d{2}/\d{2}/\d{4})'},
    {'BATCH FECHA FIN': r'BATCH FECHA FIN\s*(\d{2}/\d{2}/\d{4})'},
    {'AGUA': r'AGUA\s*([\d,]+)\s*\[% L/Vol\]'},
    {'SÓLIDOS': r'SÓLIDOS\s*([\d,]+)\s*\[% L/Vol\]'},
    {'SALES': r'SALES\s*([\d,]+)\s*\[g/m³\]'},
    {'API': r'API Seco-Seco @ 60 °F\s*([\d,.]+)\s*\[°API\]'},
    {'BRUTO ENTREGADO Hidratado': r'GSV Vol. Total Hidratado Entregado @ 15 °c\s*([\d,.]+)\s*\[L\]'},
    {'VOL. SECO-SECO': r'Vol. Seco-Seco @ 15 °C\s*([\d,.]+)\s*\[L\]'},
    {'DEN. HIDRATADA': r'DENS. HIDR. A 15ºC\s*([\d,.]+)\s*\[kg/m³\]'},
    {'DEN. SECO-SECO': r'Densidad Seco-Seco @ 15°C\s*([\d,.]+)\s*\[Kgr/m3\]'},
    {'ENTREGA CASE VOL NETO': r'CASE\s.*?\b(\d{6,7})\b'},
]

def get_actas_folder():
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    for i in range(1, outlook.Folders.Count + 1):
        account = outlook.Folders.Item(i)
        try:
            inbox = account.Folders.Item("Bandeja de entrada")
            actas_folder = inbox.Folders.Item(carpeta_correos)
            return actas_folder
        except:
            continue
    raise Exception(f"No se encontró la carpeta '{carpeta_correos}'.")

def download_attachments_from_outlook(subject_keyword, save_dir):
    actas_folder = get_actas_folder()
    messages = actas_folder.Items
    start_date = datetime.now() - timedelta(days=dias_atras)
    start_date_utc = start_date.astimezone(pytz.utc)  # Convertir a UTC
    start_date_str = start_date_utc.strftime("%d/%m/%Y %H:%M %p")
    saved_files = []

    # Filtrar mensajes en los últimos 'dias_atras' días y con el asunto que contiene el keyword especificado
    messages = messages.Restrict("[ReceivedTime] >= '" + start_date_str + "'")
    for message in messages:
        if subject_keyword in message.Subject:
            attachments = message.Attachments
            for attachment in attachments:
                if attachment.FileName.endswith('.pdf'):
                    attachment_path = os.path.join(save_dir, attachment.FileName)
                    attachment.SaveAsFile(attachment_path)
                    saved_files.append((attachment_path, message.ReceivedTime))
                    print(f'Archivo guardado en: {attachment_path}')

    if not saved_files:
        print(f"No se encontraron archivos PDF en los correos con el asunto especificado en los últimos {dias_atras} días.")
    
    return saved_files

def read_pdf(file_path):
    text = ""
    try:
        with open(file_path, 'rb') as file:
            reader = PyPDF2.PdfFileReader(file)
            num_pages = reader.numPages
            for page_num in range(num_pages):
                page = reader.getPage(page_num)
                text += page.extract_text()
    except PyPDF2.utils.PdfReadError as e:
        print(f"Error al leer el archivo PDF: {e}")
    except Exception as e:
        print(f"Se produjo un error: {e}")
    return text

def extract_data_from_text(text, prompt_list):
    data = {}
    for prompt in prompt_list:
        for key, value in prompt.items():
            pattern = re.compile(value, re.IGNORECASE)
            match = pattern.search(text)
            if match:
                extracted_value = match.group(1).strip()
                # Reemplazar la coma por el punto y el punto por la coma
                extracted_value = extracted_value.replace(',', 'TEMP').replace('.', ',').replace('TEMP', '.')
                if key == 'BATCH FECHA INICIO' or key == 'BATCH FECHA FIN':
                    extracted_value = extracted_value[:10]  # Tomar solo los primeros 10 caracteres
                data[key] = extracted_value
    return data

def extract_numbers_from_filename(filename):
    match = re.search(r'\d{3}', filename)
    if match:
        return f"ODV-{match.group(0)}"
    return 'Unknown'

def send_email_with_attachment(to, subject, body, attachment_path):
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To = to
    mail.Subject = subject
    mail.Body = body
    mail.Attachments.Add(attachment_path)
    mail.Send()
    print(f"Correo enviado a {to} con el archivo adjunto {attachment_path}")

# Descargar los archivos adjuntos
pdf_files = download_attachments_from_outlook(subject_keyword, save_dir)

# Leer y procesar los archivos PDF encontrados y almacenarlos en un DataFrame
data = []
file_numbers = []
if pdf_files:
    for pdf_path, received_time in pdf_files:
        pdf_text = read_pdf(pdf_path)
        extracted_data = extract_data_from_text(pdf_text, PROMPT_LIST)
        extracted_data['RECEPCIÓN MAIL ACTA'] = received_time.astimezone(pytz.utc).replace(tzinfo=None)
        data.append(extracted_data)
        file_numbers.append(extract_numbers_from_filename(os.path.basename(pdf_path)))
    
    df = pd.DataFrame(data)
    
    # Transponer el DataFrame
    df_transposed = df.transpose()

    # Crear una lista de nombres de columnas con los prefijos "ODV-" y los números extraídos
    new_columns = [f"{num}" for num in file_numbers]

    # Verificar si la longitud de los nuevos nombres de columnas coincide con el número de columnas del DataFrame transpuesto
    if len(new_columns) == len(df_transposed.columns):
        df_transposed.columns = new_columns
    else:
        print("Error: La longitud de los números extraídos no coincide con el número de columnas del DataFrame.")
    
    # Guardar el DataFrame transpuesto en un archivo Excel
    excel_path = os.path.join(save_dir, f"Data_PDF_Actas {fecha_formateada}.xlsx")
    df_transposed.to_excel(excel_path, header=True, index=True)

    # Ajustar el ancho de las columnas para que se adapten a los datos
    wb = load_workbook(excel_path)
    ws = wb.active
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 8)
        ws.column_dimensions[column].width = adjusted_width
    wb.save(excel_path)
    print(f'DataFrame guardado en: {excel_path}')
    
    # Enviar el correo con el archivo adjunto
    email_subject = f"Data Actas ODV Centenario {file_numbers}"
    email_body = "Adjunto se encuentra el archivo Excel con los datos procesados. Los mismos le permitiran cargar las actas en Zafiro"
    send_email_with_attachment(correo_destino, email_subject, email_body, excel_path)
    
else:
    print("No se encontraron archivos PDF para procesar.")
