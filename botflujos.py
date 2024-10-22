from exchangelib import DELEGATE, Account, Credentials, Configuration, FileAttachment
import os
import zipfile

# Configuración de conexión
creds = Credentials(username='xxxxxxx@empresa.com.ar', password='*********'
config = Configuration(server='outlook.office365.com', credentials=creds)

# Conexión a la cuenta
account = Account(primary_smtp_address='xxxxxxx@empresa.com.ar', config=config, autodiscover=True, access_type=DELEGATE)

# Seleccionar la carpeta de la bandeja de entrada
inbox = account.inbox

# Buscar los correos electrónicos que tienen archivos adjuntos y el asunto "imágenes satelitales"
emails = inbox.filter(has_attachments=True, subject__contains='imágenes satelitales')

# Iterar a través de los correos electrónicos y descargar los archivos adjuntos filtrados
attachments_folder = os.path.join(os.getcwd(), 'attachments')
if not os.path.exists(attachments_folder):
    os.makedirs(attachments_folder)

for email in emails:
    for attachment in email.attachments:
        if isinstance(attachment, FileAttachment):
            if attachment.name.endswith('.zip'):
                zip_path = os.path.join(attachments_folder, attachment.name)
                with open(zip_path, 'wb') as f:
                    f.write(attachment.content)
                with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                    zip_ref.extractall(attachments_folder)
                    os.remove(zip_path)

# Crear una vista web que muestre los datos adjuntos descargados
from flask import Flask, render_template, send_from_directory

app = Flask(__name__)

@app.route('/attachments')
def attachments():
    return render_template('attachments.html', attachments_folder=attachments_folder)

@app.route('/download_attachment/<filename>')
def download_attachment(filename):
    return send_from_directory(directory=attachments_folder, filename=filename)

if __name__ == '__main__':
    app.run(debug=True)

# Crear un archivo de Excel con los datos adjuntos descomprimidos descargados
import pandas as pd

excel_data = []
for root, dirs, files in os.walk(attachments_folder):
    for file in files:
        if file.endswith('.csv'):
            file_path = os.path.join(root, file)
            with open(file_path, 'r') as f:
                content = f.read()
                excel_data.append(content)

df = pd.DataFrame({'data': excel_data})
excel_file_path = os.path.join(os.getcwd(), 'excel_file.xlsx')
df.to_excel(excel_file_path, index=False)

