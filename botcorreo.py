#!/usr/bin/env python3
# -*- coding: utf-8 -*-
#BotCorreo - Aixa Benitez - 16/8/21

import poplib
import email
from email.header import decode_header
import os
from getpass import getpass
from email.parser import Parser
from email.policy import default
from argparse import ArgumentParser
# Conexion POP3 Outlook
mail = poplib.POP3_SSL('outlook.office365.com')
mail.user('xxxxx@outlook.es')
mail.pass_('*******')
# Selecciono la casilla de entrada
# Inbox
print("\nMensajes totales en el buzon\n")
mensajes = len(mail.list()[1])
print (mensajes)
mens=[]
for respuesta in range(mensajes):
 # Obtener el contenido
 raw_email = b"\n".join(mail.retr(respuesta+1)[1])
 mensajes = email.message_from_bytes(raw_email)
 # convertir a string
 mens.append(mensajes)
 # de donde viene el correo
 from_ = mensajes.get("From")
 subject_=mensajes.get("Subject")
 print ("\n Entrada del mensaje \n")
 print("Subject:", subject_)
 print("From:", from_)
if 'texto' in str(mensajes):
 print('palabra encontrada')
else:
 print ('palabra no encontrada')
# # correo html
 if mensajes.is_multipart():
    # Recorrer las partes del correo
    for part in mensajes.walk():
    # Extraer el contenido
     content_type = part.get_content_type()
     content_disposition = str(part.get("Content-Disposition"))
    try:
    # el cuerpo del correo
      body = part.get_payload(decode=True).decode()
    except:
     pass
    if content_type == "text/plain" and "attachment" not in content_disposition:
    # Mostrar el cuerpo del correo
       print(body)
#Eliminamos buzon
emaileliminado = len(mail.list()[1])
emaileliminado = mail.retr(emaileliminado)
print (emaileliminado)
mail.dele(emaileliminado)
mail.quit()