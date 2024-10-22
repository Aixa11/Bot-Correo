#Extraer archivos adjuntos de Outlook 

from email.message import Message
from os import path
import os
from pyexpat.errors import messages


def save_attachments(subject_prefix): # nombre del parametro
    messages.Sort("[ReceivedTime]", True)  # ordenar por fecha de recepción: de más reciente a más antigua 
    for message in Message:
        if message.Subject.startswith(subject_prefix): # test modificado
            print("saving attachments for:", message.Subject)
            for attachment in message.Attachments:
                print(attachment.FileName)
                attachment.SaveAsFile(os.path.join(path, str(attachment.FileName)))  # cambiado a nombre-archivo
            return  # salir despues del mensaje


message = messages.GetLast()
while message:  # comprobar si existe un (último) mensaje
    
    
#============== Usando Mapi ===============#
    
from aspose.email.mapi import MapiMessage

dataDir = "Data/"
downloadsDir = "Data/downloads/"
         
# Load Outlook email
message = MapiMessage.from_file(dataDir + "EmailWithAttachments.msg")

# Loop through attachments and save them
for attachment in message.attachments:
  
    # Save attachment
    attachment.save(downloadsDir + attachment.file_name)
    print("Saved...")
    
#============== Usando POP3 ===============#

# obtener mensajes

import poplib
from email.Parser import Parser

# Se establece conexion con el servidor pop de gmail
#m = poplib.POP3_SSL('pop.gmail.com',995)
#m.user('usuario@gmail.com')
#m.pass_('password')

# Se obtiene el numero de mensajes pendientes y se hace un
# bucle para cada mensaje
numero = len(m.list()[1])
for i in range (numero):
   print ("Mensaje numero"+str(i+1))
   print ("--------------------")
   # Se lee el mensaje
   response, headerLines, bytes = m.retr(i+1)
   # Se mete todo el mensaje en un unico string
   mensaje='\n'.join(headerLines)
   # Se parsea el mensaje
   p = Parser()
   email = p.parsestr(mensaje)
   # Se sacan por pantalla los campos from, to y subject
   print ("From: "+email["From"])
   print ("To: "+email["To"])
   print ("Subject: "+email["Subject"])
   # Si es un mensaje compuesto
   if (email.is_multipart()):
      # bucle para cada parte del mensaje
      for part in email.get_payload():
	 # Se mira el mime type de la parte
         tipo =  part.get_content_type()
if ("text/plain" == tipo):
		 # Si es texto plano, se saca por pantalla
		 print part.get_payload(decode=True)
if ("image/gif" == tipo):
# Si es imagen, se extrae el nombre del fichero
# adjunto y se guarda la imagen
   nombre_fichero = part.get_filename()
fp = open('recibido_'+nombre_fichero,'wb')
fp.write(part.get_payload(decode=True))
fp.close()
  else:
  # Si es mensaje simple
  tipo = email.get_content_type()
    if ("text/plain" == tipo):
	# Si es texto plano, se escribe en pantalla
        print email.get_payload(decode=True)

# Cierre de la conexion
m.quit()