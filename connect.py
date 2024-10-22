#============= Conectarse a un servidor IMAP ===========#

from aspose.email import ImapClient, SecurityOptions

# Crear e Inicializar IMAP client
client = ImapClient("imap.domain.com", 993, "user@domain.com", "pwd")

# Configurar las opciones de seguridad
client.security_options = SecurityOptions.SSLIMPLICIT

#----- Obtener los Mensajes -------#

from aspose.email import ImapClient

# Establecer una conexión con el servidor IMAP
with ImapClient("imap.gmail.com", 993, "username", "password") as conn:

    # Seleccionar carpeta
    conn.select_folder("Inbox")

    # Lista de Mensajes
    for msg in conn.list_messages():

        # Guardar mensaje
        conn.save_message(msg.unique_id, msg.unique_id + "_out.eml")
        
#================ Conexion POP3 ===================#

# Se establece conexion con el servidor pop de gmail
import poplib
m = poplib.POP3_SSL('pop.gmail.com',995)
m.user('usuario@gmail.com')
m.pass_('password')

#Establecer la conexión con smtp.gmail.com para enviar mensajes
import smtplib
mailServer = smtplib.SMTP('smtp.gmail.com',587)
mailServer.ehlo()
mailServer.starttls()
mailServer.ehlo()
mailServer.login("usuario@gmail.com","password")