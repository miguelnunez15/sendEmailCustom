# -*- coding: utf-8 -*-

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from dotenv import load_dotenv
import os
from datetime import datetime
import pandas as pd
from rich.console import Console
from rich.prompt import Prompt

# Cargar variables de entorno
load_dotenv()

html_content = ""
html_content_original = ""

from_address = os.getenv('FROM_ADDRESS')
# to_address = os.getenv('TO_ADDRESS')

smtp_server = os.getenv('SMTP_SERVER')
smtp_port = int(os.getenv('SMTP_PORT'))  # Convertir a entero
smtp_username = os.getenv('SMTP_USERNAME')
smtp_password = os.getenv('SMTP_PASSWORD')

def replaceHtmContent(row):
    global html_content

    for key in row.keys():
        if f"[[{key}]]" in html_content:
            html_content = html_content.replace(f"[[{key}]]", str(row[key]))
        

def enviar_excel(file_name, server):
    global html_content, html_content_original
    try:
        df = pd.read_excel(file_name)

        print("¿Cuál es el asunto del correo?")
        subject = input()

        print("Se va a enviar el email a las direcciones en el archivo Excel con el asunto: " + subject)

        print('¿Desea enviar el correo? (s/n)')
        answer = input()

        if answer.lower() == 's':
            os.makedirs('Output', exist_ok=True)  # Asegurar directorio Output
            log_file_path = f"Output/{datetime.now().strftime('%Y-%m-%d-%H-%M-%S')}.txt"
            
            with open(log_file_path, 'a', encoding='utf-8') as log_file:
                for index, row in df.iterrows():
                    try:
                        print(f"Enviando correo a {row['Email']}...")

                        replaceHtmContent(row)

                        message = MIMEMultipart()
                        message['From'] = from_address
                        message['To'] = row['Email']
                        message['Subject'] = subject
                        message.attach(MIMEText(html_content, 'html'))

                        server.sendmail(from_address, row['Email'], message.as_string())
                        print(f"Correo enviado a {row['Email']}")
                        log_file.write(f"Correo enviado a {row['Email']}\n")

                        html_content = html_content_original

                    except Exception as e:
                        print(f"Error al enviar el correo a {row['Email']}")
                        log_file.write(f"Error al enviar el correo a {row['Email']} - {e}\n")
    except Exception as e:
        print("Error procesando el archivo Excel:", e)

def main():
    global html_content, html_content_original
    console = Console()

    # Cargar el contenido del correo desde un archivo HTML
    try:
        with open('index.html', 'r', encoding='utf-8') as html_file:
            html_content = html_file.read()
            html_content_original = html_content
    except FileNotFoundError:
        print("El archivo index.html no se encontró.")
        return

    # Conexión al servidor SMTP
    with smtplib.SMTP(smtp_server, smtp_port) as server:
        try:
            server.starttls()
            server.login(smtp_username, smtp_password)

            # Mostrar menú
            console.print("[bold cyan]Selecciona una opción:[/bold cyan]")
            console.print("[green]1[/green] - Enviar manualmente")
            console.print("[green]2[/green] - Enviar mediante Excel [Email]")

            # Elegir opción
            opcion = Prompt.ask("Elige una opción", choices=["1", "2"])

            if opcion == "1":
                console.print("[bold green]Has seleccionado: Enviar manualmente[/bold green]")
                print('¿A quién desea enviar el correo? (direcciones separadas por comas)')
                manual_to_address = input()

                print("¿Cuál es el asunto del correo?")
                subject = input()

                print(f"Se va a enviar el email a {manual_to_address} con el asunto {subject}")

                print('¿Desea enviar el correo? (s/n)')
                answer = input()
                if answer.lower() == 's':
                    os.makedirs('Output', exist_ok=True)
                    log_file_path = f"Output/{datetime.now().strftime('%Y-%m-%d-%H-%M-%S')}.txt"

                    with open(log_file_path, 'a', encoding='utf-8') as log_file:
                        for address in manual_to_address.split(','):
                            try:
                                message = MIMEMultipart()
                                message['From'] = from_address
                                message['To'] = address.strip()
                                message['Subject'] = subject
                                message.attach(MIMEText(html_content, 'html'))

                                print(f"Enviando correo a {address.strip()}...")
                                server.sendmail(from_address, address.strip(), message.as_string())
                                log_file.write(f"Correo enviado a {address.strip()}\n")
                            except Exception as e:
                                print(f"Error al enviar el correo a {address.strip()}")
                                log_file.write(f"Error al enviar el correo a {address.strip()} - {e}\n")
                    print("Correo enviado correctamente.")
                else:
                    print("Correo no enviado.")

            elif opcion == "2":
                console.print("[bold green]Has seleccionado: Enviar mediante Excel [Email][/bold green]")
                enviar_excel('input.xlsx', server)

        except Exception as e:
            print("Error conectando al servidor SMTP:", e)

if __name__ == "__main__":
    main()
