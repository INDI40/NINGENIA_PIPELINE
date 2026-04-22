#!/usr/bin/env python3
"""
Pipeline Email Sync Server v1.1
--------------------------------
Servidor local que conecta el pipeline con tu correo IMAP/SMTP.
Ejecutar con: python email_sync.py
Mantén la ventana abierta mientras usas el pipeline.
"""
import imaplib
import smtplib
import ssl
import email as email_lib
from email.header import decode_header
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import json
import http.server
import os
import re
import urllib.request
import urllib.error
from datetime import datetime, timedelta

# ── CONFIGURACIÓN IMAP (lectura) ───────────────────────
IMAP_HOST = '217.116.0.237'
IMAP_PORT = 143
EMAIL_USER = 'david.garrido@indi40.com'
PORT       = 8765
DATA_FILE  = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'processed_emails.json')

# ── CONTRASEÑAS (fichero local, no en GitHub) ──────────
try:
    from passwords import IMAP_PASSWORD, SMTP_PASSWORDS
except ImportError:
    print('⚠️  No se encontró passwords.py. Crea el fichero con tus contraseñas.')
    IMAP_PASSWORD  = ''
    SMTP_PASSWORDS = {'ningenia': '', 'loncheria': ''}

# ── CONFIGURACIÓN SMTP (envío) ─────────────────────────
SMTP_CONFIG = {
    'ningenia': {
        'host':     '217.116.0.228',
        'port':     25,
        'user':     'david.garrido@indi40.com',
        'password': SMTP_PASSWORDS.get('ningenia', ''),
        'use_ssl':  False,
    },
    'loncheria': {
        'host':     'authsmtp.securemail.pro',
        'port':     465,
        'user':     'comercial@laloncheriadeljamon.com',
        'password': SMTP_PASSWORDS.get('loncheria', ''),
        'use_ssl':  True,
    },
}
# ──────────────────────────────────────────────────────


def load_processed():
    if os.path.exists(DATA_FILE):
        try:
            with open(DATA_FILE, encoding='utf-8') as f:
                return set(json.load(f))
        except Exception:
            pass
    return set()


def save_processed(ids):
    with open(DATA_FILE, 'w', encoding='utf-8') as f:
        json.dump(sorted(ids), f, ensure_ascii=False)


def decode_str(value):
    if value is None:
        return ''
    parts = decode_header(value)
    result = []
    for part, enc in parts:
        if isinstance(part, bytes):
            result.append(part.decode(enc or 'utf-8', errors='replace'))
        else:
            result.append(str(part))
    return ''.join(result)


def get_body(msg):
    """Extrae el texto del email (prefiere plain text, fallback a HTML limpio)."""
    body = ''
    if msg.is_multipart():
        for part in msg.walk():
            ct  = part.get_content_type()
            cd  = str(part.get('Content-Disposition', ''))
            if ct == 'text/plain' and 'attachment' not in cd:
                try:
                    charset = part.get_content_charset() or 'utf-8'
                    body = part.get_payload(decode=True).decode(charset, errors='replace')
                    break
                except Exception:
                    pass
        if not body:
            for part in msg.walk():
                if part.get_content_type() == 'text/html':
                    try:
                        charset = part.get_content_charset() or 'utf-8'
                        html = part.get_payload(decode=True).decode(charset, errors='replace')
                        body = re.sub(r'<[^>]+>', ' ', html)
                        body = re.sub(r'\s+', ' ', body).strip()
                        break
                    except Exception:
                        pass
    else:
        try:
            charset = msg.get_content_charset() or 'utf-8'
            body = msg.get_payload(decode=True).decode(charset, errors='replace')
        except Exception:
            body = ''
    return body[:3000]


def send_email(negocio, to_addr, subject, body):
    """Envía un email via SMTP según el negocio indicado. Devuelve None si OK, mensaje de error si falla."""
    cfg = SMTP_CONFIG.get(negocio)
    if not cfg:
        return f'Negocio desconocido: {negocio}. Negocios válidos: {list(SMTP_CONFIG.keys())}'
    try:
        msg = MIMEMultipart('alternative')
        msg['From']    = cfg['user']
        msg['To']      = to_addr
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain', 'utf-8'))

        if cfg['use_ssl']:
            context = ssl.create_default_context()
            with smtplib.SMTP_SSL(cfg['host'], cfg['port'], context=context) as server:
                server.login(cfg['user'], cfg['password'])
                server.send_message(msg)
        else:
            with smtplib.SMTP(cfg['host'], cfg['port'], timeout=20) as server:
                try:
                    server.login(cfg['user'], cfg['password'])
                except smtplib.SMTPNotSupportedError:
                    pass  # Algunos servidores en puerto 25 no requieren AUTH
                server.send_message(msg)
        return None
    except smtplib.SMTPAuthenticationError:
        return 'Error de autenticación SMTP: usuario o contraseña incorrectos'
    except smtplib.SMTPException as e:
        return f'Error SMTP: {e}'
    except OSError as e:
        return f'No se puede conectar al servidor SMTP ({cfg["host"]}:{cfg["port"]}): {e}'
    except Exception as e:
        return f'Error inesperado al enviar: {e}'


def fetch_emails(password, days=14):
    """Conecta al IMAP y descarga emails de los últimos `days` días."""
    try:
        mail = imaplib.IMAP4(IMAP_HOST, IMAP_PORT)
        mail.login(EMAIL_USER, password)
        mail.select('INBOX')

        since_date = (datetime.now() - timedelta(days=days)).strftime('%d-%b-%Y')
        _, data = mail.search(None, f'(SINCE {since_date})')

        emails = []
        ids = data[0].split() if data[0] else []

        for num in ids[-60:]:  # máximo 60 emails
            try:
                _, msg_data = mail.fetch(num, '(RFC822)')
                raw = msg_data[0][1]
                msg = email_lib.message_from_bytes(raw)

                msg_id   = msg.get('Message-ID', num.decode()).strip()
                subject  = decode_str(msg.get('Subject', '(Sin asunto)'))
                from_raw = decode_str(msg.get('From', ''))
                date_str = msg.get('Date', '')
                body     = get_body(msg)

                emails.append({
                    'uid':      num.decode(),
                    'msg_id':   msg_id,
                    'subject':  subject,
                    'from':     from_raw,
                    'date':     date_str,
                    'body':     body,
                })
            except Exception:
                continue

        mail.close()
        mail.logout()
        return emails, None

    except imaplib.IMAP4.error as e:
        return [], f'Error IMAP (usuario/contraseña incorrectos?): {e}'
    except OSError as e:
        return [], f'No se puede conectar al servidor IMAP ({IMAP_HOST}:{IMAP_PORT}): {e}'
    except Exception as e:
        return [], f'Error inesperado: {e}'


def call_openai(api_key, messages):
    payload = json.dumps({
        'model': 'gpt-4o-mini',
        'max_tokens': 2500,
        'response_format': {'type': 'json_object'},
        'messages': messages,
    }, ensure_ascii=False).encode('utf-8')

    req = urllib.request.Request(
        'https://api.openai.com/v1/chat/completions',
        data=payload,
        headers={
            'Content-Type':  'application/json',
            'Authorization': f'Bearer {api_key}',
        },
        method='POST'
    )
    try:
        with urllib.request.urlopen(req, timeout=45) as resp:
            result = json.loads(resp.read())
        return result['choices'][0]['message']['content'], None
    except urllib.error.HTTPError as e:
        raw = e.read().decode('utf-8', errors='replace')
        try:
            msg = json.loads(raw).get('error', {}).get('message', raw)
        except Exception:
            msg = raw[:300]
        return None, msg
    except Exception as e:
        return None, str(e)


def analyze_emails(api_key, emails, prospects):
    """Usa la IA para relacionar emails con prospectos y sugerir acciones."""

    prosp_list = [
        {
            'id':       p.get('id'),
            'empresa':  p.get('empresa', ''),
            'sector':   p.get('sector', ''),
            'contacto': p.get('contacto', ''),
            'email':    p.get('email', ''),
            'etapa':    p.get('etapa', ''),
            'negocio':  p.get('negocio', 'ningenia'),
        }
        for p in prospects
        if p.get('etapa') not in ('cerrado', 'descartado')
    ]

    email_list = [
        {
            'uid':     e['uid'],
            'from':    e['from'],
            'subject': e['subject'],
            'date':    e['date'],
            'body':    e['body'][:1000],
        }
        for e in emails
    ]

    # Etapas válidas por negocio para que la IA elija correctamente
    etapas_info = {
        'ningenia':  ['prospecto', 'contacto', 'cualificado', 'visita', 'propuesta', 'negociacion', 'cerrado'],
        'loncheria': ['prospecto', 'degustacion', 'propuesta', 'pedido', 'cerrado'],
    }

    system = (
        "Eres un asistente comercial experto. Analiza emails recibidos y relaciónalos con prospectos "
        "de un pipeline de ventas. Usa el dominio del remitente, el nombre de la empresa en el asunto "
        "o cuerpo, y el nombre del contacto para identificar coincidencias. Responde SOLO con JSON."
    )

    user = f"""PROSPECTOS ACTIVOS EN EL PIPELINE:
{json.dumps(prosp_list, ensure_ascii=False, indent=2)}

ETAPAS VÁLIDAS POR NEGOCIO:
{json.dumps(etapas_info, ensure_ascii=False)}

EMAILS RECIBIDOS:
{json.dumps(email_list, ensure_ascii=False, indent=2)}

Para cada email relacionado con un prospecto (confianza >= 6/10), devuelve:
{{
  "resultados": [
    {{
      "email_uid": "uid del email",
      "prospecto_id": "id exacto del prospecto",
      "prospecto_empresa": "nombre de la empresa",
      "confianza": 8,
      "asunto": "asunto del email",
      "remitente": "Nombre Apellido <email@dominio.com>",
      "fecha": "fecha legible en español (ej: 20 abr 2026)",
      "resumen_historial": "Resumen ejecutivo del email para el comercial: qué dice, qué implica, tono del cliente (máx 250 caracteres)",
      "nueva_etapa": "id de etapa válida para ese negocio O null si no hay cambio claro",
      "proxima_accion": "acción muy concreta que debe hacer el comercial (ej: Llamar a Juan para confirmar visita del martes)",
      "fecha_accion": "YYYY-MM-DD O null"
    }}
  ]
}}

Si ningún email coincide con confianza suficiente, devuelve {{"resultados": []}}."""

    content, err = call_openai(api_key, [
        {'role': 'system', 'content': system},
        {'role': 'user',   'content': user},
    ])

    if err:
        return None, err

    try:
        parsed = json.loads(content)
        return parsed.get('resultados', []), None
    except Exception as e:
        return None, f'JSON inválido en respuesta IA: {e}'


# ══════════════════════════════════════════════════════
# HTTP SERVER
# ══════════════════════════════════════════════════════

class SyncHandler(http.server.BaseHTTPRequestHandler):

    def do_OPTIONS(self):
        self.send_response(200)
        self._cors_headers()
        self.end_headers()

    def do_GET(self):
        if self.path == '/ping':
            self._respond(200, {'status': 'ok', 'server': 'Pipeline Email Sync v1.0'})
        else:
            self._respond(404, {'error': 'Not found'})

    def do_POST(self):
        if self.path not in ('/sync', '/send'):
            self._respond(404, {'error': 'Not found'})
            return

        # Leer body
        try:
            length = int(self.headers.get('Content-Length', 0))
            body   = json.loads(self.rfile.read(length).decode('utf-8'))
        except Exception as e:
            self._respond(400, {'error': f'Petición inválida: {e}'})
            return

        # ── /send ─────────────────────────────────────────
        if self.path == '/send':
            negocio = body.get('negocio', '').strip()
            to_addr = body.get('to', '').strip()
            subject = body.get('subject', '').strip()
            text    = body.get('body', '').strip()
            if not to_addr:
                self._respond(400, {'error': 'Falta el destinatario (to)'})
                return
            if not subject:
                self._respond(400, {'error': 'Falta el asunto (subject)'})
                return
            print(f'  → Enviando email a {to_addr} | Negocio: {negocio}')
            err = send_email(negocio, to_addr, subject, text)
            if err:
                print(f'  ✗ Error SMTP: {err}')
                self._respond(500, {'error': err})
            else:
                print(f'  ✓ Email enviado correctamente')
                self._respond(200, {'ok': True, 'msg': f'Email enviado a {to_addr}'})
            return

        # ── /sync ─────────────────────────────────────────
        api_key    = body.get('openai_key', '').strip()
        email_pass = body.get('email_pass', '').strip()
        prospects  = body.get('prospects', [])

        if not email_pass:
            self._respond(400, {'error': 'Falta la contraseña del correo. Configúrala en ⚙ Configuración.'})
            return
        if not api_key:
            self._respond(400, {'error': 'Falta la OpenAI API Key. Configúrala en ⚙ Configuración.'})
            return
        if not prospects:
            self._respond(200, {'updates': [], 'checked': 0, 'new': 0, 'msg': 'Sin prospectos activos'})
            return

        print(f'  → Conectando al IMAP...')
        emails, err = fetch_emails(email_pass)
        if err:
            self._respond(500, {'error': err})
            return
        print(f'  → {len(emails)} emails encontrados en los últimos 14 días')

        # Filtrar ya procesados
        processed  = load_processed()
        new_emails = [e for e in emails if e['msg_id'] not in processed]
        print(f'  → {len(new_emails)} emails nuevos (sin procesar)')

        if not new_emails:
            self._respond(200, {
                'updates': [],
                'checked': len(emails),
                'new':     0,
                'msg':     f'Sin emails nuevos ({len(emails)} revisados)',
            })
            return

        print(f'  → Analizando con IA...')
        updates, err = analyze_emails(api_key, new_emails, prospects)
        if err:
            self._respond(500, {'error': f'Error IA: {err}'})
            return

        # Marcar todos los nuevos como procesados
        processed.update(e['msg_id'] for e in new_emails)
        save_processed(processed)

        # Enriquecer actualizaciones con snippet del cuerpo
        email_map = {e['uid']: e for e in new_emails}
        for u in (updates or []):
            orig = email_map.get(u.get('email_uid'), {})
            u['body_snippet'] = orig.get('body', '')[:400]

        matched = len(updates or [])
        print(f'  → {matched} emails relacionados con prospectos')

        self._respond(200, {
            'updates': updates or [],
            'checked': len(emails),
            'new':     len(new_emails),
            'matched': matched,
            'msg':     f'{len(new_emails)} emails nuevos, {matched} relacionados con prospectos',
        })

    def _cors_headers(self):
        self.send_header('Access-Control-Allow-Origin',  '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')

    def _respond(self, code, data):
        body = json.dumps(data, ensure_ascii=False).encode('utf-8')
        self.send_response(code)
        self._cors_headers()
        self.send_header('Content-Type',   'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, fmt, *args):
        ts = datetime.now().strftime('%H:%M:%S')
        print(f'[{ts}] {fmt % args}')


# ══════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════

if __name__ == '__main__':
    print()
    print('╔══════════════════════════════════════════════╗')
    print('║      Pipeline Email Sync Server v1.0         ║')
    print('╠══════════════════════════════════════════════╣')
    print(f'║  Escuchando en: http://localhost:{PORT}         ║')
    print(f'║  Correo:  {EMAIL_USER}     ║')
    print(f'║  IMAP:    {IMAP_HOST}:{IMAP_PORT}          ║')
    print('╠══════════════════════════════════════════════╣')
    print('║  ✅ Listo. Abre el pipeline en el navegador. ║')
    print('║  Cierra esta ventana para detener.           ║')
    print('╚══════════════════════════════════════════════╝')
    print()

    try:
        server = http.server.ThreadingHTTPServer(('127.0.0.1', PORT), SyncHandler)
        server.serve_forever()
    except KeyboardInterrupt:
        print('\n[Servidor detenido]')
        server.server_close()
    except OSError as e:
        if 'Address already in use' in str(e) or '10048' in str(e):
            print(f'\n⚠️  El puerto {PORT} ya está en uso. El servidor ya está corriendo.')
            input('Pulsa Enter para cerrar...')
        else:
            raise
