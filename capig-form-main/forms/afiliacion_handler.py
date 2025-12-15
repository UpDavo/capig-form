import os
from typing import Dict, List

from django.conf import settings
from gspread.utils import rowcol_to_a1

from capig_form.services.google_sheets_service import get_google_sheet

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
    "NO.",
    "NRO",
    "NUM",
    "NUMERO",
}

TRANSLITERATION = str.maketrans(
    {
        "\u00c1": "A",
        "\u00c9": "E",
        "\u00cd": "I",
        "\u00d3": "O",
        "\u00da": "U",
        "\u00dc": "U",
        "\u00d1": "N",
        "\u00c0": "A",
        "\u00c8": "E",
        "\u00cc": "I",
        "\u00d2": "O",
        "\u00d9": "U",
        "\u00c2": "A",
        "\u00ca": "E",
        "\u00ce": "I",
        "\u00d4": "O",
        "\u00db": "U",
        "\u00b0": "",
    }
)

def _normalize(col: str) -> str:
    """Normaliza el nombre de columna para comparar."""
    col = col.strip().upper().replace(" ", "_")
    col = col.translate(TRANSLITERATION)
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
        elif key in {"GENERO", "GENERO"}:
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
        elif key in ALT_ID_KEYS:
            filas.append(data.get("_id_autoinc", ""))
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

    # Buscar columna de ID y calcular el siguiente correlativo
    id_col_index = None
    for idx, col in enumerate(header):
        if _normalize(col) in ALT_ID_KEYS:
            id_col_index = idx + 1  # 1-based index para gspread
            break

    next_id = ""
    if id_col_index:
        col_values = sheet.col_values(id_col_index)
        # Saltar titulo (fila 1) y encabezado (fila 2)
        existing_ids: List[int] = []
        for raw in col_values[2:]:
            try:
                existing_ids.append(int(str(raw).strip()))
            except (TypeError, ValueError):
                continue
        next_id = str(max(existing_ids) + 1) if existing_ids else "1"
    data["_id_autoinc"] = next_id

    fila = _build_fila(header, data)

    # Primera fila libre (incluye encabezados); no usar filas personalizadas ni duplicadas
    next_row = len(sheet.get_all_values()) + 1

    start_cell = rowcol_to_a1(next_row, 1)
    end_cell = rowcol_to_a1(next_row, len(header))
    sheet.update(f"{start_cell}:{end_cell}", [fila])
    return True
