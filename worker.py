import os
import sys
import re
import csv
import json
import time
import email
import gspread
import imaplib
import requests
from colorama import init, Fore
from email.header import decode_header
from datetime import datetime, timedelta, timezone
from google.oauth2.service_account import Credentials
from dotenv import load_dotenv

# Cargar variables de entorno desde archivo .env local si existe
load_dotenv()

init(autoreset=True)

# ====================================================================================================================================================================================================================
# ── CONFIGURACIONES (SE CARGAN DE LEY DE ENV O DE COPIA LOCAL) ─────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────
# ====================================================================================================================================================================================================================

GMAIL_USER               = os.getenv("GMAIL_USER")
GMAIL_APP_PASS           = os.getenv("GMAIL_APP_PASS")
IMAP_SERVER              = os.getenv("IMAP_SERVER", "imap.gmail.com")
SUBJECT_FILTER           = os.getenv("SUBJECT_FILTER", "DealerOn")
CREDS_PATH               = os.getenv("CREDS_PATH", "credentials.json")
SPREADSHEET_ID           = os.getenv("GOOGLE_SHEET_ID")
SHEET_NAME               = os.getenv("SHEET_NAME", "Sheet1")
WHATSAPP_ACCESS_TOKEN    = os.getenv("META_ACCESS_TOKEN")
WHATSAPP_PHONE_NUMBER_ID = os.getenv("META_PHONE_NUMBER_ID")
WHATSAPP_TEMPLATE_NAME   = os.getenv("WHATSAPP_TEMPLATE_NAME", "dealeron_lead_notification")

# Receptores de notificaciones de WhatsApp (se pueden separar por comas en la variable)
recipients_env = os.getenv("WHATSAPP_RECIPIENTS")
if recipients_env:
    WHATSAPP_RECIPIENTS = [num.strip() for num in recipients_env.split(",") if num.strip()]
else:
    WHATSAPP_RECIPIENTS = []

OUTPUT_CSV = os.path.join(os.path.dirname(os.path.abspath(__file__)), "dealeron_leads.csv")

# ====================================================================================================================================================================================================================

COLUMNS = {
    "FechaRecibido":      "Fecha Recibido",
    "EnviadoPor":         "Enviado Por",
    "FechaEnvio":         "Fecha de Envío",
    "Nombre":             "Nombre",
    "Apellido":           "Apellido",
    "Telefono":           "Teléfono",
    "CorreoElectronico":  "Correo Electrónico",
    "AnioInteres":        "Año de Interés",
    "MarcaInteres":       "Marca de Interés",
    "ModeloInteres":      "Modelo de Interés",
}
KEYS_ORDER   = list(COLUMNS.keys())
HEADERS_ES   = list(COLUMNS.values())

# ====================================================================================================================================================================================================================

def decode_str(s):
    if s is None:
        return ""
    parts = decode_header(s)
    result = []
    for part, enc in parts:
        if isinstance(part, bytes):
            result.append(part.decode(enc or "utf-8", errors="replace"))
        else:
            result.append(part)
    return "".join(result)

def get_text_body(msg):
    body = ""
    if msg.is_multipart():
        for part in msg.walk():
            ct   = part.get_content_type()
            disp = str(part.get("Content-Disposition") or "")
            if ct == "text/plain" and "attachment" not in disp:
                charset = part.get_content_charset() or "utf-8"
                body = part.get_payload(decode=True).decode(charset, errors="replace")
                break
    else:
        charset = msg.get_content_charset() or "utf-8"
        body = msg.get_payload(decode=True).decode(charset, errors="replace")
    return body

def extract_field(body, field_name):
    pattern = rf"^{re.escape(field_name)}\s*:\s*(.+)$"
    match = re.search(pattern, body, re.MULTILINE | re.IGNORECASE)
    return match.group(1).strip() if match else ""

def parse_comments_field(body):
    result = {"EstadoCompra": "", "ComentariosAdicionales": "", "ConsentimientoTCPA": ""}
    raw = extract_field(body, "Comments")
    if not raw:
        return result
    m = re.search(r"Shopping Status\s*:\s*([^,]+)", raw, re.IGNORECASE)
    if m:
        result["EstadoCompra"] = m.group(1).strip()
    m = re.search(r"Additional Comments\s*:\s*([^,]+)", raw, re.IGNORECASE)
    if m:
        result["ComentariosAdicionales"] = m.group(1).strip()
    m = re.search(r"TCPA_Consent_Text\s*:\s*(.+)", raw, re.IGNORECASE)
    if m:
        result["ConsentimientoTCPA"] = m.group(1).strip()
    return result

def parse_submitted_line(body):
    pattern = r"(.+?)\s*\(.+?\)\s+on\s+(\d{1,2}/\d{1,2}/\d{4}\s+\d{1,2}:\d{2}:\d{2}\s+(?:AM|PM))"
    match = re.search(pattern, body, re.IGNORECASE)
    if match:
        return match.group(1).strip(), match.group(2).strip()
    return "", ""

def parse_email_body(body, received_date):
    submitted_by, submitted_date = parse_submitted_line(body)
    comments = parse_comments_field(body)
    row = {
        "FechaRecibido":          received_date,
        "EnviadoPor":             submitted_by,
        "FechaEnvio":             submitted_date,
        "Nombre":                 extract_field(body, "First Name"),
        "Apellido":               extract_field(body, "Last Name"),
        "Telefono":               extract_field(body, "Daytime Phone Number"),
        "CorreoElectronico":      extract_field(body, "Email Address"),
        "AnioInteres":            extract_field(body, "Interested Year"),
        "MarcaInteres":           extract_field(body, "Interested Make"),
        "ModeloInteres":          extract_field(body, "Interested Model"),
        "VersionInteres":         extract_field(body, "Interested Trim Level"),
        "EstadoCompra":           comments["EstadoCompra"],
        "ComentariosAdicionales": comments["ComentariosAdicionales"],
        "ConsentimientoTCPA":     comments["ConsentimientoTCPA"],
        "EsFormularioPlataforma": extract_field(body, "IsPlatformForm"),
        "ApellidoMaterno":        extract_field(body, "MothersLastName"),
        "OpcionRecibirInfo":      extract_field(body, "ReceiveInfoOption"),
    }
    return row

def build_imap_date_range(first_run=True):
    today = datetime.now().date()
    if first_run:
        if today.weekday() == 0:  # Lunes, buscar desde el sábado (hace 2 días)
            since_date = today - timedelta(days=2)
        else:
            since_date = today - timedelta(days=1)
    else:
        since_date = today
    since_str  = since_date.strftime("%d-%b-%Y")
    before_str = (today + timedelta(days=1)).strftime("%d-%b-%Y")
    return since_str, before_str

def get_sheet():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds_json = os.getenv("GOOGLE_CREDENTIALS_JSON")
    if creds_json:
        creds = Credentials.from_service_account_info(json.loads(creds_json), scopes=scopes)
    else:
        if os.path.exists(CREDS_PATH):
            creds = Credentials.from_service_account_file(CREDS_PATH, scopes=scopes)
        else:
            raise FileNotFoundError("No se encontraron credenciales de Google (ni en GOOGLE_CREDENTIALS_JSON ni en ruta local).")
            
    client = gspread.authorize(creds)
    sh     = client.open_by_key(SPREADSHEET_ID)
    try:
        ws = sh.worksheet(SHEET_NAME)
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=SHEET_NAME, rows=1000, cols=len(KEYS_ORDER))
        print(Fore.YELLOW + f"Pestaña '{SHEET_NAME}' creada.")
    return ws

def normalize_datetime(val):
    val = val.strip()
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %I:%M:%S %p", "%d/%m/%Y %H:%M:%S"):
        try:
            dt = datetime.strptime(val, fmt)
            return dt.strftime("%Y-%m-%d %H:%M:%S")
        except ValueError:
            continue
    match = re.match(r"(\d{4}-\d{2}-\d{2})\s+(\d{1,2}):(\d{2}):(\d{2})", val)
    if match:
        return f"{match.group(1)} {int(match.group(2)):02d}:{match.group(3)}:{match.group(4)}"
    return val

def send_whatsapp_alerts(new_leads):
    if not WHATSAPP_ACCESS_TOKEN or not WHATSAPP_PHONE_NUMBER_ID or not WHATSAPP_RECIPIENTS:
        return
    url = f"https://graph.facebook.com/v20.0/{WHATSAPP_PHONE_NUMBER_ID}/messages"
    headers = {
        "Authorization": f"Bearer {WHATSAPP_ACCESS_TOKEN}",
        "Content-Type": "application/json"
    }
    for lead in new_leads:
        nombre_completo = f"{lead.get('Nombre', '')} {lead.get('Apellido', '')}".strip()
        if not nombre_completo:
            nombre_completo = "Cliente Sin Nombre"
        telefono = lead.get("Telefono", "") or "No proporcionado"
        correo = lead.get("CorreoElectronico", "") or "No proporcionado"
        marca = lead.get("MarcaInteres", "") or "N/A"
        modelo = lead.get("ModeloInteres", "") or "N/A"
        parameters = [
            {"type": "text", "text": nombre_completo},
            {"type": "text", "text": telefono},
            {"type": "text", "text": correo},
            {"type": "text", "text": marca},
            {"type": "text", "text": modelo}
        ]
        payload = {
            "messaging_product": "whatsapp",
            "type": "template",
            "template": {
                "name": WHATSAPP_TEMPLATE_NAME,
                "language": {
                    "code": "es_MX"
                },
                "components": [
                    {
                        "type": "body",
                        "parameters": parameters
                    }
                ]
            }
        }
        for recipient in WHATSAPP_RECIPIENTS:
            num_clean = recipient.strip().replace("+", "").replace(" ", "")
            if len(num_clean) == 10:
                num_clean = f"52{num_clean}"
            payload_to_send = payload.copy()
            payload_to_send["to"] = num_clean
            print(Fore.CYAN + f"[WhatsApp] Enviando alerta a {num_clean} para el lead {nombre_completo}...")
            try:
                res = requests.post(url, headers=headers, json=payload_to_send)
                res_data = res.json()
                if res.status_code == 200:
                    print(Fore.GREEN + f"[WhatsApp] Mensaje enviado exitosamente a {num_clean}.")
                else:
                    print(Fore.RED + f"[WhatsApp] Error al enviar mensaje a {num_clean} (Codigo {res.status_code}):")
                    print(Fore.RED + json.dumps(res_data, indent=2, ensure_ascii=True))
            except Exception as e:
                print(Fore.RED + f"[WhatsApp] Error de conexion al enviar a {num_clean}: {e}")

def send_whatsapp_to_clients(new_leads):
    if not WHATSAPP_ACCESS_TOKEN or not WHATSAPP_PHONE_NUMBER_ID:
        return
    url = f"https://graph.facebook.com/v20.0/{WHATSAPP_PHONE_NUMBER_ID}/messages"
    headers = {
        "Authorization": f"Bearer {WHATSAPP_ACCESS_TOKEN}",
        "Content-Type": "application/json"
    }
    for lead in new_leads:
        telefono = lead.get("Telefono", "").strip()
        if not telefono or telefono.lower() == "no proporcionado":
            continue
        num_clean = "".join(filter(str.isdigit, telefono))
        if len(num_clean) == 10:
            num_clean = f"52{num_clean}"
            
        nombre = lead.get("Nombre", "").strip() or "Cliente"
        marca = lead.get("MarcaInteres", "").strip() or "N/A"
        modelo = lead.get("ModeloInteres", "").strip() or "N/A"
        anio = str(lead.get("AnioInteres", "")).strip() or "N/A"
        
        parameters = [
            {"type": "text", "text": nombre},
            {"type": "text", "text": marca},
            {"type": "text", "text": modelo},
            {"type": "text", "text": anio}
        ]
        
        payload = {
            "messaging_product": "whatsapp",
            "to": num_clean,
            "type": "template",
            "template": {
                "name": "solicitud_cotizacion_auto",
                "language": {"code": "es_MX"},
                "components": [{"type": "body", "parameters": parameters}]
            }
        }
        
        print(Fore.CYAN + f"[WhatsApp Cliente] Enviando mensaje de bienvenida a {num_clean} ({nombre})...")
        try:
            res = requests.post(url, headers=headers, json=payload)
            res_data = res.json()
            if res.status_code == 200:
                print(Fore.GREEN + f"[WhatsApp Cliente] Mensaje enviado exitosamente a {num_clean}.")
            else:
                print(Fore.RED + f"[WhatsApp Cliente] Error al enviar mensaje a {num_clean} (Codigo {res.status_code}):")
                print(Fore.RED + json.dumps(res_data, indent=2, ensure_ascii=True))
        except Exception as e:
            print(Fore.RED + f"[WhatsApp Cliente] Error de conexion al enviar a {num_clean}: {e}")

def write_to_sheets(rows):
    ws = get_sheet()
    existing = ws.get_all_values()
    filled_existing = [r for r in existing if any(cell.strip() for cell in r[:10])]
    existing_keys = set()
    if not filled_existing:
        print(Fore.YELLOW + f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] [ADVERTENCIA] La hoja está completamente vacía. Escribiendo desde A2.")
        next_row = 2
    else:
        try:
            idx_fecha = filled_existing[0].index("Fecha Recibido") if "Fecha Recibido" in filled_existing[0] else 0
            idx_correo = filled_existing[0].index("Correo Electrónico") if "Correo Electrónico" in filled_existing[0] else 6
            for r in filled_existing[1:]:
                if len(r) > max(idx_fecha, idx_correo):
                    fecha_norm = normalize_datetime(r[idx_fecha])
                    existing_keys.add((fecha_norm, r[idx_correo].strip()))
        except Exception:
            for r in filled_existing[1:]:
                if len(r) > 6:
                    fecha_norm = normalize_datetime(r[0])
                    existing_keys.add((fecha_norm, r[6].strip()))
        next_row = len(filled_existing) + 1
    data_rows = []
    new_leads = []
    for row in rows:
        local_fecha = normalize_datetime(row.get("FechaRecibido", ""))
        key = (local_fecha, row.get("CorreoElectronico", "").strip())
        if key in existing_keys:
            continue
        data_rows.append([row.get(k, "") for k in KEYS_ORDER])
        new_leads.append(row)
    if data_rows:
        total_needed = next_row + len(data_rows) - 1
        if total_needed > ws.row_count:
            rows_to_add = total_needed - ws.row_count
            ws.add_rows(rows_to_add)
        range_name = f"A{next_row}"
        ws.update(range_name=range_name, values=data_rows, value_input_option="USER_ENTERED")
        print(Fore.GREEN + f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] [Sheets] {len(data_rows)} registro(s) nuevo(s) agregado(s).")
        send_whatsapp_alerts(new_leads)
        send_whatsapp_to_clients(new_leads)
    else:
        print(Fore.YELLOW + f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] [Sheets] No hay registros nuevos para agregar (todos ya existen).")

def run_sync(first_run=True):
    since_str, before_str = build_imap_date_range(first_run)
    now_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    print(Fore.CYAN + f"[{now_str}] [Gmail] Buscando correos del {since_str} al {before_str}...")
    sys.stdout.flush()
    try:
        mail = imaplib.IMAP4_SSL(IMAP_SERVER)
        mail.login(GMAIL_USER, GMAIL_APP_PASS)
    except imaplib.IMAP4.error as e:
        print(Fore.RED + f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] [ERROR] No se pudo conectar a Gmail: {e}")
        return False
    mail.select("INBOX")
    search_criteria = (f'(SUBJECT "{SUBJECT_FILTER}" SINCE "{since_str}" BEFORE "{before_str}")')
    status, data = mail.search(None, search_criteria)
    if status != "OK" or not data[0]:
        mail.logout()
        print(Fore.YELLOW + f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] [Gmail] No se encontraron correos nuevos.")
        return True
    msg_ids = data[0].split()
    rows = []
    for num, msg_id in enumerate(msg_ids, start=1):
        status, msg_data = mail.fetch(msg_id, "(RFC822)")
        if status != "OK":
            print(Fore.RED + f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] [{num}] Error al obtener correo ID {msg_id.decode()}")
            continue
        msg = email.message_from_bytes(msg_data[0][1])
        date_header = msg.get("Date", "")
        try:
            parsed_date   = email.utils.parsedate_to_datetime(date_header)
            tz_hermosillo = timezone(timedelta(hours=-7))
            received_date = parsed_date.astimezone(tz_hermosillo).strftime("%Y-%m-%d %H:%M:%S")
        except Exception:
            received_date = date_header
        body = get_text_body(msg)
        submission_url = extract_field(body, "Submission URL")
        if "seminuevo" in submission_url.lower():
            continue
        row  = parse_email_body(body, received_date)
        rows.append(row)
    mail.logout()
    if not rows:
        print(Fore.YELLOW + f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] [Gmail] No se encontraron leads válidos (excluidos/seminuevos).")
        return True
    with open(OUTPUT_CSV, mode="w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=KEYS_ORDER, extrasaction="ignore")
        writer.writerow(COLUMNS)
        writer.writerows(rows)
    write_to_sheets(rows)
    return True

def is_active_time(dt):
    weekday = dt.weekday()
    if weekday == 6:  # Domingo no se ejecuta
        return False
    t = dt.time()
    w1_start = datetime.strptime("08:30", "%H:%M").time()
    w1_end = datetime.strptime("13:30", "%H:%M").time()
    if weekday == 5:  # Sábado solo primer turno (8:30 AM a 1:30 PM)
        return w1_start <= t < w1_end
    # Lunes a Viernes
    w2_start = datetime.strptime("15:30", "%H:%M").time()
    w2_end = datetime.strptime("19:00", "%H:%M").time()
    return (w1_start <= t < w1_end) or (w2_start <= t < w2_end)

def main():
    # Validar variables de entorno requeridas
    missing_vars = []
    if not GMAIL_USER:
        missing_vars.append("GMAIL_USER")
    if not GMAIL_APP_PASS:
        missing_vars.append("GMAIL_APP_PASS")
    if not SPREADSHEET_ID:
        missing_vars.append("GOOGLE_SHEET_ID")
    
    # Validar credenciales de Google
    creds_json = os.getenv("GOOGLE_CREDENTIALS_JSON")
    if not creds_json and not (CREDS_PATH and os.path.exists(CREDS_PATH)):
        missing_vars.append("GOOGLE_CREDENTIALS_JSON (o archivo de credenciales local en CREDS_PATH)")
        
    if missing_vars:
        print(Fore.RED + "\n[ERROR] Faltan variables de entorno obligatorias para iniciar el worker:")
        for var in missing_vars:
            print(Fore.RED + f"  - {var}")
        print(Fore.YELLOW + "\nPor favor, configura estas variables en tu panel de Railway (o en un archivo .env local) para poder iniciar.")
        sys.exit(1)

    print(Fore.CYAN + "\n=== GmailLeadsCSV — Extractor de Leads DealerOn (Modo Continuo) ===")
    print(Fore.CYAN + f"Cuenta  : {GMAIL_USER}")
    print(Fore.CYAN + f"Asunto  : {SUBJECT_FILTER}")
    sys.stdout.flush()
    
    first_run = True
    was_active = None
    
    while True:
        try:
            now = datetime.now()
            active = is_active_time(now)
            
            if active:
                if was_active is not True:
                    print(Fore.GREEN + f"\n[{now.strftime('%Y-%m-%d %H:%M:%S')}] >>> ENTRANDO EN PERIODO ACTIVO DE EJECUCIÓN <<<")
                    first_run = True
                    was_active = True
                
                success = run_sync(first_run=first_run)
                if success:
                    first_run = False
            else:
                if was_active is not False:
                    print(Fore.YELLOW + f"\n[{now.strftime('%Y-%m-%d %H:%M:%S')}] >>> ENTRANDO EN MODO DE REPOSO <<<")
                    was_active = False
                
            sys.stdout.flush()
        except Exception as e:
            print(Fore.RED + f"[EXCEPCIÓN] Error en ciclo continuo: {e}")
            sys.stdout.flush()
            
        # Calcular los segundos restantes para alinearse con el segundo 00 del siguiente minuto
        now = datetime.now()
        sleep_duration = 60.0 - now.second - (now.microsecond / 1000000.0) + 0.1
        time.sleep(sleep_duration)

if __name__ == "__main__":
    main()
