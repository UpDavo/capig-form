import logging
import os
import re
import unicodedata
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

PHONE_COLUMN_SEQUENCE = [
    "TELEFONO_EMPRESA_1",
    "TELEFONO_EMPRESA_2",
    "CONTACTO",
    "CONTACTO_2",
    "CONTACTO_4",
    "CONTACTO_6",
    "CONTACTO_8",
]
EMAIL_COLUMN_SEQUENCE = [
    "EMAIL",
    "EMAIL2",
    "COREO_ELSCTRONICO",
    "COREO_ELSCTRONICO3",
    "COREO_ELSCTRONICO5",
    "CORREO",
    "COREO_ELSCTRONICO9",
]
PHONE_OVERRIDE_KEYS = {
    "CONTACTO": "gerente_general_contacto",
    "CONTACTO_2": "gerente_fin_contacto",
    "CONTACTO_4": "gerente_rrhh_contacto",
    "CONTACTO_6": "gerente_comercial_contacto",
    "CONTACTO_8": "gerente_produccion_contacto",
}
PHONE_COLUMN_ALIASES = {
    "TELEFONO_EMPRESA_1": ["TELEFONO_EMPRESA", "TELEFONO"],
    "TELEFONO_EMPRESA_2": ["TELEFONO_2"],
}
EMAIL_OVERRIDE_KEYS = {
    "COREO_ELSCTRONICO": "gerente_general_email",
    "COREO_ELSCTRONICO3": "gerente_fin_email",
    "COREO_ELSCTRONICO5": "gerente_rrhh_email",
    "CORREO": "gerente_comercial_email",
    "COREO_ELSCTRONICO9": "gerente_produccion_email",
}
EMAIL_COLUMN_ALIASES = {
    "EMAIL2": ["EMAIL_2"],
    "COREO_ELSCTRONICO": ["CORREO_ELECTRONICO"],
    "COREO_ELSCTRONICO3": ["CORREO_ELECTRONICO3", "CORREO_ELECTRONICO_3"],
    "COREO_ELSCTRONICO5": ["CORREO_ELECTRONICO5", "CORREO_ELECTRONICO_5"],
    "COREO_ELSCTRONICO9": ["CORREO_ELECTRONICO9", "CORREO_ELECTRONICO_9"],
}

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
    "ACTIVIDAD": "actividad",
    "NOMBRE_REP_LEGAL": "representante",
    "CARGO": "cargo",
    "NOMBRE_GTE_GRAL": "gerente_general_nombre",
    "GENERO": "genero",
    "NOMBRE_GTE_FIN": "gerente_fin_nombre",
    "NOMBRE_GTE_RRHH": "gerente_rrhh_nombre",
    "GERENTE_COMERC": "gerente_comercial_nombre",
    "GERENTE_COMERCIAL": "gerente_comercial_nombre",
    "GERENTE_PRODUCCCION": "gerente_produccion_nombre",
    "GERENTE_PRODUCCION": "gerente_produccion_nombre",
    "SECTOR": "sector",
    "TAMANO": "tamano",
    "TAMANO_EMPRESA": "tamano",
    "ESTADO": "estado",
}

COLABORADORES_KEYS = {"NO_COLABORADORES", "COLABORADORES", "NUM_COLABORADORES", "NUMERO_COLABORADORES"}


def _normalize(col: str) -> str:
    """Normaliza el nombre de columna para comparar de forma tolerante."""
    text = unicodedata.normalize("NFD", str(col or "").strip().upper())
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
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


def _clean_list_values(values) -> List[str]:
    cleaned = []
    for value in values or []:
        txt = str(value or "").strip()
        if txt:
            cleaned.append(txt)
    return cleaned


def _build_sequence_assignments(payload: Dict[str, str], sequence, overrides, list_field: str) -> Dict[str, str]:
    assigned: Dict[str, str] = {}

    for column_key, data_key in overrides.items():
        value = str(payload.get(data_key, "") or "").strip()
        if value:
            assigned[column_key] = value

    list_values = _clean_list_values(payload.get(list_field, []))
    next_index = 0
    for column_key in sequence:
        if column_key in assigned:
            continue
        if next_index < len(list_values):
            assigned[column_key] = list_values[next_index]
            next_index += 1
        else:
            assigned[column_key] = ""
    return assigned


def _expand_assignment_aliases(assignments: Dict[str, str], alias_map) -> Dict[str, str]:
    expanded = dict(assignments)
    for canonical_key, aliases in alias_map.items():
        value = assignments.get(canonical_key, "")
        for alias in aliases:
            expanded[alias] = value
    return expanded


def _build_fila(header: List[str], data: Dict[str, str]) -> List[str]:
    """
    Construye una fila con la misma cantidad de columnas que la hoja,
    llenando solo las columnas mapeadas y el resto con "".
    """
    fila = []
    payload = dict(data)
    payload["ruc"] = limpiar_ruc(payload.get("ruc", ""))
    phone_assignments = _build_sequence_assignments(
        payload,
        PHONE_COLUMN_SEQUENCE,
        PHONE_OVERRIDE_KEYS,
        "telefonos",
    )
    phone_assignments = _expand_assignment_aliases(phone_assignments, PHONE_COLUMN_ALIASES)
    email_assignments = _build_sequence_assignments(
        payload,
        EMAIL_COLUMN_SEQUENCE,
        EMAIL_OVERRIDE_KEYS,
        "emails",
    )
    email_assignments = _expand_assignment_aliases(email_assignments, EMAIL_COLUMN_ALIASES)

    for col in header:
        key = _normalize(col)
        if key in phone_assignments:
            fila.append(phone_assignments.get(key, ""))
        elif key in email_assignments:
            fila.append(email_assignments.get(key, ""))
        elif key in COLUMN_TO_DATA_KEY:
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
