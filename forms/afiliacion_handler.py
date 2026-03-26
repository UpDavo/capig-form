import logging
import os
import re
from typing import Dict, List

from django.conf import settings
from gspread.utils import rowcol_to_a1

from capig_form.services.google_sheets_service import (
    ensure_row_capacity,
    find_first_empty_row,
    get_google_sheet,
)
from forms.utils import limpiar_ruc

logger = logging.getLogger(__name__)

ALT_ID_KEYS = {
    "ID_UNICO",
    "ID_INTERNO",
    "ID",
    "ID_SOCIO",
    "CODIGO_SOCIO",
    "CODIGO",
    "CLAVE",
    "CLAVE_UNICA",
    "NO",
    "NRO",
    "NUM",
    "NUMERO",
}

COLUMN_TO_DATA_KEY = {
    "RAZON_SOCIAL": "razon_social",
    "RUC": "ruc",
    "FECHA_AFILIACION": "fecha_afiliacion",
    "CIUDAD": "ciudad",
    "DIRECCION": "direccion",
    "TELEFONO_EMPRESA_1": "telefono_1",
    "TELEFONO_EMPRESA": "telefono_1",
    "TELEFONO": "telefono_1",
    "TELEFONO_EMPRESA_2": "telefono_2",
    "TELEFONO_2": "telefono_2",
    "EMAIL": "email_1",
    "EMAIL2": "email_2",
    "EMAIL_2": "email_2",
    "ACTIVIDAD": "actividad",
    "NOMBRE_REP_LEGAL": "representante",
    "CARGO": "cargo",
    "NOMBRE_GTE_GRAL": "gerente_general_nombre",
    "GENERO": "genero",
    "CONTACTO": "gerente_general_contacto",
    "CORREO_ELECTRONICO": "gerente_general_email",
    "COREO_ELSCTRONICO": "gerente_general_email",
    "NOMBRE_GTE_FIN": "gerente_fin_nombre",
    "CONTACTO_2": "gerente_fin_contacto",
    "CORREO_ELECTRONICO3": "gerente_fin_email",
    "CORREO_ELECTRONICO_3": "gerente_fin_email",
    "COREO_ELSCTRONICO3": "gerente_fin_email",
    "NOMBRE_GTE_RRHH": "gerente_rrhh_nombre",
    "CONTACTO_4": "gerente_rrhh_contacto",
    "CORREO_ELECTRONICO5": "gerente_rrhh_email",
    "CORREO_ELECTRONICO_5": "gerente_rrhh_email",
    "COREO_ELSCTRONICO5": "gerente_rrhh_email",
    "GERENTE_COMERC": "gerente_comercial_nombre",
    "GERENTE_COMERCIAL": "gerente_comercial_nombre",
    "CONTACTO_6": "gerente_comercial_contacto",
    "CORREO": "gerente_comercial_email",
    "GERENTE_PRODUCCCION": "gerente_produccion_nombre",
    "GERENTE_PRODUCCION": "gerente_produccion_nombre",
    "CONTACTO_8": "gerente_produccion_contacto",
    "CORREO_ELECTRONICO9": "gerente_produccion_email",
    "CORREO_ELECTRONICO_9": "gerente_produccion_email",
    "COREO_ELSCTRONICO9": "gerente_produccion_email",
    "SECTOR": "sector",
    "TAMANO": "tamano",
    "TAMANO_EMPRESA": "tamano",
    "ESTADO": "estado",
}

COLABORADORES_KEYS = {
    "NO_COLABORADORES",
    "COLABORADORES",
    "NUM_COLABORADORES",
    "NUMERO_COLABORADORES",
}

TRANSLITERATION = str.maketrans(
    {
        "Á": "A",
        "É": "E",
        "Í": "I",
        "Ó": "O",
        "Ú": "U",
        "Ü": "U",
        "Ñ": "N",
        "À": "A",
        "È": "E",
        "Ì": "I",
        "Ò": "O",
        "Ù": "U",
        "Â": "A",
        "Ê": "E",
        "Î": "I",
        "Ô": "O",
        "Û": "U",
        "º": "",
        "°": "",
    }
)


def _normalize(col: str) -> str:
    """Normaliza el nombre de columna para comparar de forma tolerante."""
    text = str(col or "").strip().upper()
    text = text.translate(TRANSLITERATION)
    text = re.sub(r"[^A-Z0-9]+", "_", text)
    text = re.sub(r"_+", "_", text).strip("_")
    return text


def _get_header_row(sheet):
    """Prefiere la fila 2 como encabezado real de SOCIOS y usa fila 1 como fallback."""
    fallback = None
    for row_number in (2, 1):
        header = sheet.row_values(row_number)
        if not any((cell or "").strip() for cell in header):
            continue
        header_keys = {_normalize(cell) for cell in header if (cell or "").strip()}
        if {"RUC", "RAZON_SOCIAL", "FECHA_AFILIACION"} & header_keys:
            return row_number, header
        if fallback is None:
            fallback = (row_number, header)
    if fallback is not None:
        return fallback
    raise RuntimeError("La hoja SOCIOS no tiene encabezados disponibles.")


def _next_autoincrement_value(sheet, header_row: int, header: List[str]) -> str:
    """Calcula el siguiente correlativo si la hoja tiene una columna tipo ID/No."""
    id_col_index = None
    for idx, col in enumerate(header, start=1):
        if _normalize(col) in ALT_ID_KEYS:
            id_col_index = idx
            break

    if not id_col_index:
        return ""

    existing_ids: List[int] = []
    for raw in sheet.col_values(id_col_index)[header_row:]:
        try:
            existing_ids.append(int(str(raw).strip()))
        except (TypeError, ValueError):
            continue
    return str(max(existing_ids) + 1) if existing_ids else "1"


def _build_fila(header: List[str], data: Dict[str, str]) -> List[str]:
    """
    Construye una fila con la misma cantidad de columnas que la hoja,
    llenando solo las columnas mapeadas y el resto con "".
    """
    fila = []
    payload = dict(data)
    payload["ruc"] = limpiar_ruc(payload.get("ruc", ""))

    for col in header:
        key = _normalize(col)
        if key in COLUMN_TO_DATA_KEY:
            fila.append(payload.get(COLUMN_TO_DATA_KEY[key], ""))
        elif key in COLABORADORES_KEYS:
            fila.append(payload.get("colaboradores", ""))
        elif key in ALT_ID_KEYS:
            fila.append(payload.get("_id_autoinc", ""))
        else:
            fila.append("")
    return fila


def guardar_nuevo_afiliado_en_google_sheets(data: Dict[str, str]) -> bool:
    """
    Guarda un nuevo registro de afiliado en la hoja SOCIOS,
    alineado con los encabezados reales de la hoja.
    """
    sheet_id = os.getenv("SHEET_PATH") or getattr(settings, "SHEET_PATH", "")
    if not sheet_id:
        raise RuntimeError("SHEET_PATH no esta configurado.")

    sheet = get_google_sheet(sheet_id, "SOCIOS")
    header_row, header = _get_header_row(sheet)

    payload = dict(data)
    payload["_id_autoinc"] = _next_autoincrement_value(sheet, header_row, header)

    fila = _build_fila(header, payload)
    if len(fila) < len(header):
        fila += [""] * (len(header) - len(fila))
    elif len(fila) > len(header):
        fila = fila[: len(header)]

    next_row = find_first_empty_row(sheet, start_row=header_row + 1)
    ensure_row_capacity(sheet, next_row)
    start_cell = rowcol_to_a1(next_row, 1)
    end_cell = rowcol_to_a1(next_row, len(header))
    try:
        sheet.update(f"{start_cell}:{end_cell}", [fila], value_input_option="USER_ENTERED")
    except Exception:  # pragma: no cover - depende de API externa
        logger.exception("Error al insertar afiliado en SOCIOS fila %s", next_row)
        raise
    return True
