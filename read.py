#leer los correos
# Antes instalar --> pip install Aspose.Email-for-Python-via-NET

from email.message import EmailMessage

dataDir = "Data/"

# Crear una instancia de MailMessage cargando un archivo Eml
message = EmailMessage.load(dataDir + "test.eml")

# Obtener la información del remitente, la información del destinatario, el asunto, el cuerpo html y el cuerpo del texto 
print("Sender: " + str(message.from_address))

for receiver in enumerate(message.to):
    print("Receiver: " + receiver)

print("Subject: " + message.subject)

print("HtmlBody: " + message.html_body)

print("TextBody: " + message.body)

#Extraer texto sin formato del cuerpo HTML del correo electrónico

dataDir = "Data/"

# Cree una instancia de MailMessage cargando un archivo Eml
message = EmailMessage.load(dataDir + "test.eml")

# Obtener texto del cuerpo HTML 
print("HTML body text: " + message.get_html_body_text(False))

#Leer encabezados de un correo

dataDir = "Data/"

# Cree una instancia de MailMessage cargando un archivo EML
message = EmailMessage.load(dataDir + "email-headers.eml");
print("\n\nheaders:\n\n")

# Imprime todos los encabezados
index = 0
for index, header in enumerate(message.headers):
    print(header + " - ", end=" ")
    print (message.headers.get(index))

#======== Leer correos con POP3 =========#

# n° de mensajes sin leer
import math as m
numero = len(m.list()[1])

for i in range (numero):
   print ("Mensaje numero"+str(i+1))
   print ("--------------------")
   # Se lee el mensaje
   response, headerLines, bytes = m.retr(i+1)