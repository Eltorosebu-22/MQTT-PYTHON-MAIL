import smtplib
import openpyxl
import ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import rsa
from cryptography.fernet import Fernet
import paho.mqtt.client as mqtt
from urllib import parse, request
import time
import math
import requests

############
contadorcorreo=0
val1=float(0.0)
val1_2=float(0.0)
torque_list=[] ## lista de los 10 torques detectados
celdasexcel_list=[] ##lista que posiciona los torques en las celdas adecuadas
cont = 2  ###contador celdas en letras
cont1 = 2 ###contador desplazamiento vertical
tor=50
mandar=0  ## variable que permite el envio de datos
recibo=0 ##variable que dice si recibio un dato de torque
###Private key#################################################
with open("private.pem", "rb") as f:
    private_key = rsa.PrivateKey.load_pkcs1(f.read())

encryptedm = open("PSW", "rb").read()
clear = rsa.decrypt(encryptedm, private_key)
##################################################
x = 1
#######################Mandado de datos via gmail###############
smtp_port = 587  # Puerto standard de gmail SMTP
smtp_server = "smtp.gmail.com"  # Servidor google SMTP

email_from = "lsendmessagetoro@gmail.com"
email_list = ["luisgarna09@gmail.com","agus_als_ver@hotmail.com"]  # lista de correos que reciben la información

pswd = (clear.decode())


##################################################################
def send_emails(email_list):
    for person in email_list:
        # Make the body of the email
        body = f"""
        Estimado Gerente de Control de procesos de Adient o a quien corresponda,

        Espero que este mensaje le encuentre bien. Me complace presentarle el informe de datos diario acerca del 
        rendimiento del robot cooperativo para el proceso de atornillado del asiento correspondiente al Audi Q5.
        Este informe contiene una visión general detallada de los datos clave relacionados con el desempeño y las 
        operaciones.

        Espero sea de ayuda, y quedamos al pendiente de cualquier duda o aclaración.
        """

        # make a MIME object to define parts of the email
        msg = MIMEMultipart()
        msg['From'] = email_from
        msg['To'] = person
        msg['Subject'] = "Reporte diario, del rendimiento"

        # Attach the body of the message
        msg.attach(MIMEText(body, 'plain'))

        # Define the file to attach
        filename = "BOOKOUT.xlsx"

        # Open the file in python as a binary
        attachment = open(filename, 'rb')  # r for read and b for binary

        # Encode as base 64
        attachment_package = MIMEBase('application', 'octet-stream')
        attachment_package.set_payload((attachment).read())
        encoders.encode_base64(attachment_package)
        attachment_package.add_header('Content-Disposition', "attachment; filename= " + filename)
        msg.attach(attachment_package)

        # Cast as string
        text = msg.as_string()

        # Connect with the server
        print("Connecting to server...")
        TIE_server = smtplib.SMTP(smtp_server, smtp_port)
        TIE_server.starttls()
        TIE_server.login(email_from, pswd)
        print("Succesfully connected to server")
        print()

        # Send emails to "person" as list is iterated
        print(f"Sending email to: {person}...")
        TIE_server.sendmail(email_from, person, text)
        print(f"Email sent to: {person}")
        print()

    # Close the port
    TIE_server.quit()
def excel():
    length=len(celdasexcel_list)
    n=0
    workbook = openpyxl.load_workbook("LIBROIN .xlsx")
    sheet = workbook.active
    sheet.title = "Changed"
    for c in range(length):
        r=celdasexcel_list[n]
        r1=torque_list[n]
        n=n+1
        sheet[r]=r1
    workbook.save('BOOKOUT.xlsx')

def encryption():
    with open("goblalkey.key", "rb") as mykey:
        key = mykey.read()

    f = Fernet(key)
    ###################open book to encrypt###############
    with open("BOOKOUT.xlsx", "rb") as orginal_file:
        original = orginal_file.read()

    encrypted = f.encrypt(original)
    ###############encrypt the book####################33
    with open("../../../Downloads/BOOKOUT2.xlsx", "wb") as encrypted_file:
        encrypted_file.write(encrypted)
def obtener_variable(cont):
    if cont==2 or cont==12:
        letra="B"
    elif cont==3:
        letra="C"
    elif cont==4:
        letra="D"
    elif cont==5:
        letra="E"
    elif cont==6:
        letra="F"
    elif cont==7:
        letra="G"
    elif cont==8:
        letra="H"
    elif cont==9:
        letra="I"
    elif cont==10:
        letra="J"
    elif cont==11:
        letra="K"
    return letra


def on_connect(client, userdata, flags, rc):
    print("Connected with result code " + str(rc))
    client.subscribe("tema")
    client.subscribe("/DUAL/RelevadoEsfuerzos/Home2/Temperatura/Setpoint/")


def on_message(client, userdata, msg):
    if msg.topic == "tema":
        str0 = str(msg.payload)
        global torque_lis
        torque_lis = str0.split(",")
        print(torque_lis)####entran todos los valores

def main():
    payload = build_payload(varl1, varl_2)
    # print("[INFO] Attemping to send data")
    post_request(payload)
c=0

if __name__ == '__main__':
    recibo=1

    while (True):
        if c == 0:
            mqttc = mqtt.Client()
            mqttc.on_connect = on_connect
            mqttc.on_message = on_message
            mqttc.connect("192.168.130.216", 1883, 60)
            mqttc.loop_start()
            main()
            c=1
        print("Waiting for data.............")
        valor_torque1 = torque_lis[1]
        valor_torque2 = torque_lis[2]
        valor_torque3 = torque_lis[3]
        valor_torque4 = torque_lis[4]
        valor_torque5 = torque_lis[5]
        recibo_datosexcel=torque_lis[6]
        mandar_correo=torque_lis[7]
        if recibo_datosexcel=="P":
            recibo=1
        if recibo_datosexcel == "T" and recibo==1:
            torque_listnew = []
            torque_listnew.append(valor_torque1), torque_listnew.append(valor_torque2), torque_listnew.append(valor_torque3),torque_listnew.append(valor_torque4),torque_listnew.append(valor_torque5)
            torque_list = torque_list + torque_listnew
            length_dato = len(torque_listnew)
            for i in range(length_dato):
                letra_obtenida = obtener_variable(cont)
                celda_concatenada = letra_obtenida + "" + str(cont1)
                celdasexcel_list.append(celda_concatenada)
                cont = cont + 1
            cont1 = cont1 + 1
            cont = 2
            print(torque_list)
            print(celdasexcel_list)
            recibo=0
        if mandar_correo == "E'":
            excel()
            send_emails(email_list)
            mandar = 0
            torque_list.clear()
            celdasexcel_list.clear()
        c=0

