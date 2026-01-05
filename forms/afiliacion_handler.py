import os
from typing import Dict, List

from django.conf import settings
from gspread.utils import rowcol_to_a1

from capig_form.services.google_sheets_service import (
    get_google_sheet,
    find_first_empty_row,
)


def _normalize(col: str) -> str:
    """Normaliza el nombre de columna para comparar."""
    col = col.strip().upper().replace(" ", "_")
    col = (
        col.replace("Á", "A")
        .replace("É", "E")
        .replace("Í", "I")
        .replace("Ó", "O")
        .replace("Ú", "U")
        .replace("Ñ", "N")
    )
    return col


def _build_fila(header: List[str], data: Dict[str, str]) -> List[str]:
    """
    Construye una fila con la misma cantidad de columnas que la hoja,
    llenando solo las columnas mapeadas y el resto con "".
    """
    filas = []
    for col in header:
        key = _normalize(col)
        if key == "RAZON_SOCIAL":
            filas.append(data.get("razon_social", ""))
        elif key == "RUC":
            filas.append(data.get("ruc", ""))
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
        elif key in {"GENERO", "GÉNERO", "GENERO"}:
            filas.append(data.get("genero", ""))
        elif key in {"NO._COLABORADORES", "NO_COLABORADORES", "NO.COLABORADORES", "COLABORADORES", "NUM_COLABORADORES", "NUMERO_COLABORADORES"}:
            filas.append(data.get("colaboradores", ""))

        elif key == "SECTOR":
            filas.append(data.get("sector", ""))
        elif key == "TAMANO":
            filas.append(data.get("tamano", ""))
        elif key == "ESTADO":
            filas.append(data.get("estado", ""))
        else:
            filas.append("")
    return filas


def guardar_nuevo_afiliado_en_google_sheets(data: Dict[str, str]) -> bool:
    """
    Guarda un nuevo registro de afiliado en la hoja SOCIOS,
    alineado con los encabezados originales (fila 2).
    """
    sheet_id = os.getenv("SHEET_PATH") or getattr(settings, "SHEET_PATH", "")
    if not sheet_id:
        raise RuntimeError("SHEET_PATH no esta configurado.")

    sheet = get_google_sheet(sheet_id, "SOCIOS")

    # Encabezados reales en fila 2
    header = sheet.row_values(2)
    fila = _build_fila(header, data)

    # Primera fila realmente libre (omite filas con formato pero sin datos)
    next_row = find_first_empty_row(sheet, start_row=2)
    start_cell = rowcol_to_a1(next_row, 1)
    end_cell = rowcol_to_a1(next_row, len(header))
    sheet.update(f"{start_cell}:{end_cell}", [fila])
    return True
