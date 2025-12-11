# -*- coding: utf-8 -*-
import json
import logging
import re
import traceback

import gspread
from django.conf import settings
from google.oauth2.service_account import Credentials
from gspread.exceptions import APIError, SpreadsheetNotFound, WorksheetNotFound

try:
    from googleapiclient.errors import HttpError
except ImportError:  # pragma: no cover - dependencia opcional
    HttpError = Exception


logger = logging.getLogger(__name__)

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
REQUIRED_SERVICE_FIELDS = {"private_key", "client_email", "project_id"}


def _print_utf8(message):
    print(str(message).encode("utf-8", errors="replace").decode("utf-8"))


def _load_service_account_info():
    try:
        raw_service = settings.SERVICE
    except AttributeError as exc:
        logger.exception("La variable SERVICE no esta definida en settings.")
        raise RuntimeError("SERVICE no esta configurado en settings.py") from exc

    try:
        info = json.loads(raw_service)
    except json.JSONDecodeError as exc:
        logger.exception("El valor de SERVICE no contiene JSON valido.")
        raise RuntimeError("El JSON en SERVICE esta mal formado. Revisa el .env.") from exc
    except Exception as exc:
        logger.exception("Error inesperado al interpretar SERVICE.")
        raise RuntimeError("No se pudo interpretar SERVICE del .env.") from exc

    missing = [field for field in REQUIRED_SERVICE_FIELDS if not info.get(field)]
    if missing:
        message = f"Faltan campos obligatorios en SERVICE: {missing}"
        logger.error(message)
        raise RuntimeError(message)

    return info


SERVICE_ACCOUNT_INFO = _load_service_account_info()


# ======================
# CLIENTE DE AUTENTICACION
# ======================
def _get_client():
    try:
        creds = Credentials.from_service_account_info(SERVICE_ACCOUNT_INFO, scopes=SCOPES)
        return gspread.authorize(creds)
    except Exception as exc:
        logger.exception("Error autenticando con Google Sheets.")
        _print_utf8(traceback.format_exc())
        raise RuntimeError("No fue posible autenticarse con Google Sheets.") from exc


# ==========================
# OBTENER HOJA POR NOMBRE
# ==========================
def get_google_sheet(sheet_id, worksheet_name):
    try:
        client = _get_client()
        spreadsheet = client.open_by_key(sheet_id)
        worksheets = spreadsheet.worksheets()
        available_titles = [ws.title for ws in worksheets]

        _print_utf8(f"sheet_id solicitado: {sheet_id}")
        _print_utf8(f"worksheet_name solicitado: {repr(worksheet_name)}")
        _print_utf8(f"Hojas disponibles en el documento: {available_titles}")

        if worksheet_name not in available_titles:
            logger.warning(
                "El nombre de hoja '%s' no coincide exactamente con las hojas disponibles: %s",
                worksheet_name,
                available_titles,
            )

        return spreadsheet.worksheet(worksheet_name)

    except WorksheetNotFound as exc:
        msg = f"La hoja '{worksheet_name}' no fue encontrada. Revisa mayusculas y espacios."
        logger.exception(msg)
        _print_utf8(msg)
        _print_utf8(traceback.format_exc())
        raise

    except SpreadsheetNotFound as exc:
        msg = f"No se encontro el Google Sheet con ID: {sheet_id}"
        logger.exception(msg)
        _print_utf8(msg)
        _print_utf8(traceback.format_exc())
        raise

    except Exception as exc:
        msg = f"Error inesperado al acceder a la hoja: {exc}"
        logger.exception(msg)
        _print_utf8(msg)
        _print_utf8(traceback.format_exc())
        raise


# ========================
# INSERTAR UNA FILA
# ========================
def insert_row_to_sheet(sheet_id, worksheet_name, data):
    try:
        sheet = get_google_sheet(sheet_id, worksheet_name)
        logger.info("Insertando fila en hoja '%s': len=%s datos=%s", worksheet_name, len(data), data)
        _print_utf8(f"Insertando {len(data)} valores en '{worksheet_name}': {data}")
        sheet.append_row(data)
        return True

    except (WorksheetNotFound, SpreadsheetNotFound) as exc:
        logger.exception("No se pudo insertar porque no se encontro el documento u hoja.")
        _print_utf8(traceback.format_exc())
        return False

    except (APIError, HttpError) as exc:
        logger.exception("La API de Google rechazo la insercion en '%s'.", worksheet_name)
        _print_utf8(traceback.format_exc())
        return False

    except Exception as exc:
        logger.exception("Error inesperado al insertar fila en '%s'.", worksheet_name)
        _print_utf8(traceback.format_exc())
        return False

def update_sheet_with_dataframe(sheet_id, worksheet_name, df):
    """
    Borra la hoja indicada y sube el contenido del DataFrame.
    """
    try:
        sheet = get_google_sheet(sheet_id, worksheet_name)

        # Limpiar la hoja
        sheet.clear()

        # Subir encabezados y datos
        headers = df.columns.values.tolist()
        data = df.values.tolist()
        all_data = [headers] + data

        _print_utf8(f"Subiendo {len(data)} filas + headers a '{worksheet_name}' en {sheet_id}")
        sheet.update(all_data)
        return True

    except Exception as exc:
        logger.exception("Error al actualizar hoja con DataFrame.")
        _print_utf8(traceback.format_exc())
        return False

# ========================
# LEER COLUMNA
# ========================
def get_column_data(sheet_id, worksheet_index=0, column='A', start_row=2):
    try:
        column = column.strip()
        if not column:
            raise ValueError("El parametro 'column' no puede estar vacio.")
        if not re.fullmatch(r"[A-Za-z]", column):
            raise ValueError("El parametro 'column' debe ser una sola letra de la A a la Z.")

        client = _get_client()
        sheet = client.open_by_key(sheet_id).get_worksheet(worksheet_index)

        col_num = ord(column.upper()) - ord('A') + 1
        column_values = sheet.col_values(col_num)
        result = [val.strip() for val in column_values[start_row - 1:] if val.strip()]

        logger.info("Se obtuvieron %s valores desde columna '%s'.", len(result), column)
        return result

    except ValueError as exc:
        logger.exception("Se recibio un identificador de columna invalido: '%s'.", column)
        _print_utf8(str(exc))
        _print_utf8(traceback.format_exc())
        raise

    except Exception as exc:
        logger.exception("Error al leer columna '%s'.", column)
        _print_utf8(traceback.format_exc())
        return []
