"""Envía por SMTP los PDF agregados o actualizados en una publicación."""

import mimetypes
import os
import smtplib
import ssl
from email.message import EmailMessage
from pathlib import Path


def required(name: str) -> str:
    value = os.environ.get(name, "").strip()
    if not value:
        raise RuntimeError(f"Falta configurar el secreto {name}")
    return value


host = required("SMTP_HOST")
port = int(required("SMTP_PORT"))
user = required("SMTP_USER")
password = required("SMTP_PASSWORD")
recipients = [
    address.strip()
    for address in required("DESTINATARIOS_INFORME").replace(";", ",").split(",")
    if address.strip()
]
files = [Path(line.strip()) for line in required("PDF_FILES").splitlines() if line.strip()]

missing = [str(path) for path in files if not path.is_file()]
if missing:
    raise FileNotFoundError("No se encontraron los PDF: " + ", ".join(missing))

repository = os.environ.get("REPOSITORY", "Dashboard HH Mantención SAP")
commit = os.environ.get("COMMIT_SHA", "")[:7]
message = EmailMessage()
message["Subject"] = f"Nuevo informe PDF de mantenimiento ({commit})"
message["From"] = user
message["To"] = ", ".join(recipients)
message.set_content(
    "Se publicó una nueva actualización del informe de mantenimiento.\n\n"
    f"Repositorio: {repository}\n"
    f"Actualizado por: {os.environ.get('UPDATED_BY', 'No informado')}\n"
    f"Detalle: {os.environ.get('COMMIT_MESSAGE', 'Actualización de informe')}\n"
    f"Enlace: {os.environ.get('UPDATE_URL', '')}\n\n"
    "Los archivos PDF actualizados se encuentran adjuntos a este correo."
)

for path in files:
    mime, _ = mimetypes.guess_type(path.name)
    maintype, subtype = (mime or "application/pdf").split("/", 1)
    message.add_attachment(path.read_bytes(), maintype=maintype, subtype=subtype, filename=path.name)

context = ssl.create_default_context()
if port == 465:
    with smtplib.SMTP_SSL(host, port, context=context) as smtp:
        smtp.login(user, password)
        smtp.send_message(message)
else:
    with smtplib.SMTP(host, port) as smtp:
        smtp.starttls(context=context)
        smtp.login(user, password)
        smtp.send_message(message)

print(f"Correo enviado a {len(recipients)} destinatario(s) con {len(files)} PDF.")
