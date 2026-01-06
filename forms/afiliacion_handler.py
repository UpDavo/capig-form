import logging
import os
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


def _normalize(col: str) -> str:
    """Normaliza el nombre de columna eliminando acentos y espacios."""
    col = unicodedata.normalize("NFKD", col or "").encode(
        "ascii", "ignore").decode()
    col = col.strip().upper().replace(" ", "_")
    return col


def _build_fila(header: List[str], data: Dict[str, str], seq_no: int = None) -> List[str]:
    """
    Construye una fila con la misma cantidad de columnas que la hoja,
    llenando solo las columnas mapeadas y el resto con "".
    """
    filas = []
    ruc_norm = limpiar_ruc(data.get("ruc", ""))
    for col in header:
        key = _normalize(col)
        if key in {"NO", "NO.", "Nº", "N\u00ba", "N"}:
            filas.append(seq_no if seq_no is not None else "")
        elif key == "RAZON_SOCIAL":
            filas.append(data.get("razon_social", ""))
        elif key == "RUC":
            filas.append(ruc_norm)
        elif key == "FECHA_AFILIACION":
            filas.append(data.get("fecha_afiliacion", ""))
        elif key == "CIUDAD":
            filas.append(data.get("ciudad", ""))
        elif key == "DIRECCION":
            filas.append(data.get("direccion", ""))
        elif key in {"TELEFONO_EMPRESA_1", "TELEFONO_EMPRESA", "TELEFONO"}:
            filas.append(data.get("telefono", ""))
        elif key == "EMAIL":
            filas.append(data.get("email", ""))
        elif key == "NOMBRE_REP_LEGAL":
            filas.append(data.get("representante", ""))
        elif key == "CARGO":
            filas.append(data.get("cargo", ""))
        elif key == "GENERO":
            filas.append(data.get("genero", ""))
        elif key in {
            "NO._COLABORADORES",
            "NO_COLABORADORES",
            "NO.COLABORADORES",
            "COLABORADORES",
            "NUM_COLABORADORES",
            "NUMERO_COLABORADORES",
        }:
            filas.append(data.get("colaboradores", ""))
        elif key == "SECTOR":
            filas.append(data.get("sector", ""))
        elif key in {"TAMANO", "TAMANO_EMPRESA"}:
            filas.append(data.get("tamano", ""))
        elif key == "ESTADO":
            filas.append(data.get("estado", ""))
        else:
            filas.append("")
    return filas


def guardar_nuevo_afiliado_en_google_sheets(data: Dict[str, str]) -> bool:
    """
    Guarda un nuevo registro de afiliado en la hoja SOCIOS,
    alineado con los encabezados originales.
    """
    sheet_id = os.getenv("SHEET_PATH") or getattr(settings, "SHEET_PATH", "")
    if not sheet_id:
        raise RuntimeError("SHEET_PATH no esta configurado.")

    sheet = get_google_sheet(sheet_id, "SOCIOS")

    # Encabezados reales (preferir fila 1; fallback fila 2)
    header_row = 1
    header = sheet.row_values(header_row)
    if not any((c or "").strip() for c in header):
        header_row = 2
        header = sheet.row_values(header_row)
    if not header:
        raise RuntimeError("La hoja SOCIOS no tiene encabezados (fila 1/2 vacía).")

    # Calcular correlativo (número de registro) con base en cantidad de filas de datos
    data_rows = sheet.get_all_records(
        head=header_row,
        value_render_option="UNFORMATTED_VALUE",
        numericise_ignore=["all"],
    )
    seq_no = len(data_rows) + 1

    fila = _build_fila(header, data, seq_no=seq_no)
    # Ajustar longitud de fila al header
    if len(fila) < len(header):
        fila += [""] * (len(header) - len(fila))
    elif len(fila) > len(header):
        fila = fila[: len(header)]

    # Primera fila realmente libre (omite filas con formato pero sin datos)
    # Calcular la siguiente fila usando la cantidad de registros, para evitar saltos por filas con formato
    next_row = header_row + len(data_rows) + 1
    ensure_row_capacity(sheet, next_row)
    start_cell = rowcol_to_a1(next_row, 1)
    end_cell = rowcol_to_a1(next_row, len(header))
    try:
        sheet.update(f"{start_cell}:{end_cell}", [fila])
    except Exception as exc:  # pragma: no cover - dep externa
        logger.exception("Error al insertar afiliado en SOCIOS fila %s", next_row)
        raise
    return True
