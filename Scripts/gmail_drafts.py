#!/usr/bin/env python3
"""
Módulo para crear borradores en Gmail usando la API de Google.
Requiere credentials.json en la carpeta Scripts/.
"""

import base64
import json
import os
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email import encoders
from pathlib import Path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

SCOPES = ['https://www.googleapis.com/auth/gmail.compose']

ORDEN_REGIONES = ['R1', 'R2', 'R3', 'R4', 'R5', 'R6', 'R7', 'R8', 'R9', 'MOM', 'BUI']

NOMBRES_REGION = {
    'R1':  ('R1',                   'la R1'),
    'R2':  ('R2',                   'la R2'),
    'R3':  ('R3',                   'la R3'),
    'R4':  ('R4',                   'la R4'),
    'R5':  ('R5',                   'la R5'),
    'R6':  ('R6',                   'la R6'),
    'R7':  ('R7',                   'la R7'),
    'R8':  ('R8',                   'la R8'),
    'R9':  ('R9',                   'la R9'),
    'MOM': ('Mobile Marketing',     'Mobile Marketing'),
    'BUI': ('Business Intelligence','Business Intelligence'),
}


def get_service(script_dir):
    """Obtiene el servicio Gmail autenticado. Abre el navegador la primera vez."""
    credentials_path = os.path.join(script_dir, 'credentials.json')
    token_path = os.path.join(script_dir, 'token.json')

    creds = None
    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(credentials_path, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(token_path, 'w') as f:
            f.write(creds.to_json())

    return build('gmail', 'v1', credentials=creds)


def token_valido(script_dir):
    """Devuelve True si ya hay un token guardado y válido."""
    token_path = os.path.join(script_dir, 'token.json')
    if not os.path.exists(token_path):
        return False
    try:
        creds = Credentials.from_authorized_user_file(token_path, SCOPES)
        if creds.valid:
            return True
        if creds.expired and creds.refresh_token:
            creds.refresh(Request())
            with open(token_path, 'w') as f:
                f.write(creds.to_json())
            return True
    except Exception:
        pass
    return False


def _construir_mensaje(to_list, subject, body, attachment_path=None):
    """Construye el objeto MIME del correo."""
    msg = MIMEMultipart()
    msg['to'] = ', '.join(to_list)
    msg['subject'] = subject
    msg.attach(MIMEText(body, 'plain', 'utf-8'))

    if attachment_path and os.path.exists(attachment_path):
        with open(attachment_path, 'rb') as f:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(f.read())
        encoders.encode_base64(part)
        filename = os.path.basename(attachment_path)
        part.add_header('Content-Disposition', f'attachment; filename="{filename}"')
        msg.attach(part)

    return base64.urlsafe_b64encode(msg.as_bytes()).decode()


def crear_borradores(script_dir, archivos_por_region, contraseñas, periodo_legible):
    """
    Crea 11 borradores en Gmail (uno por región).

    archivos_por_region: dict { 'R1': '/ruta/R1-leads-...xlsx', ... }
    contraseñas:         dict { 'R1': 'R1TELXXX', ... }  — solo regiones con leads
    periodo_legible:     str  'Viernes 13' o 'Viernes 13 al Domingo 15'
    """
    config_path = os.path.join(script_dir, 'config.json')
    with open(config_path, 'r', encoding='utf-8') as f:
        config = json.load(f)

    destinatarios = config.get('destinatarios', {})
    service = get_service(script_dir)

    resultados = []

    for region in ORDEN_REGIONES:
        nombre_subject, nombre_mail = NOMBRES_REGION[region]
        to_list = destinatarios.get(region, [])
        tiene_leads = region in contraseñas

        subject = f"Reporte de leads {nombre_subject} ({periodo_legible})"

        if tiene_leads:
            pwd = contraseñas[region]
            body = (
                f"Buenos días equipo.\n\n"
                f"Comparto el reporte de leads generados para {nombre_mail} "
                f"correspondiente al período del {periodo_legible}.\n"
                f"Contraseña: {pwd}\n\n"
                f"Quedo pendiente.\n"
                f"Saludos."
            )
            attachment = archivos_por_region.get(region)
        else:
            body = (
                f"Buenos días equipo.\n\n"
                f"Les comento que no se han generado nuevos leads para {nombre_mail} "
                f"del período del {periodo_legible}.\n\n"
                f"Quedo pendiente.\n"
                f"Saludos."
            )
            attachment = None

        try:
            raw = _construir_mensaje(to_list, subject, body, attachment)
            draft = service.users().drafts().create(
                userId='me',
                body={'message': {'raw': raw}}
            ).execute()
            resultados.append({
                'region': region,
                'ok': True,
                'draft_id': draft.get('id'),
                'tiene_leads': tiene_leads,
            })
        except Exception as e:
            resultados.append({
                'region': region,
                'ok': False,
                'error': str(e),
                'tiene_leads': tiene_leads,
            })

    return resultados
