"""
Construye un dataset consolidado de capacitaciones combinando:
- Hoja CAPACITACIONES (histórico con totales y valor)
- Hoja CAPACITACIONES_FINAL (nuevos registros)
- Hoja BASE DE DATOS (para mapear RUC por razón social)
- Hoja TAMANO_EMPRESA_GLOBAL (para obtener el tamaño por RUC)

Salida:
- Excel local: capig_form/services/data/capacitaciones_dash.xlsx
- (Opcional) hojas en Google Sheets: CAPACITACIONES_DASH_DATA

No toca otras lógicas ni configuraciones del sistema.
"""
import os
import unicodedata
from collections import defaultdict
from typing import Dict, List, Optional, Tuple

import pandas as pd

# Rutas
BASE_FILENAME = "datos_completos.xlsx"
OUTPUT_FILENAME = "capacitaciones_dash.xlsx"
BASE_PATH = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "..", "data", BASE_FILENAME)
)
OUTPUT_PATH = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "data", OUTPUT_FILENAME)
)


def _normalize(text: str) -> str:
    if text is None:
        txt = ""
    else:
        txt = str(text).strip().upper()
    txt = unicodedata.normalize("NFD", txt)
    txt = "".join(ch for ch in txt if unicodedata.category(ch) != "Mn")
    return txt


def _normalize_ruc(raw) -> str:
    return str(raw or "").replace("'", "").replace('"', "").strip().lstrip("0")


def _find_col(headers: List[str], needle: str) -> int:
    target = _normalize(needle)
    for idx, h in enumerate(headers):
        if target in _normalize(h):
            return idx
    return -1


def _load_base_maps(xl: pd.ExcelFile) -> Tuple[Dict[str, str], Dict[str, str]]:
    """
    Retorna:
    - name_to_ruc: map por razón social normalizada -> RUC
    - ruc_to_tam: map de RUC -> TAMANO desde TAMANO_EMPRESA_GLOBAL
    """
    name_to_ruc: Dict[str, str] = {}
    ruc_to_tam: Dict[str, str] = {}

    # BASE DE DATOS: detectar filas de encabezado y procesar bloques
    df_base = xl.parse("BASE DE DATOS", header=None).fillna("")
    rows = df_base.values.tolist()
    header_idxs = []
    for i, row in enumerate(rows):
        if _find_col(row, "RUC") >= 0 and _find_col(row, "RAZON_SOCIAL") >= 0:
            header_idxs.append(i)

    header_idxs = sorted(set(header_idxs))
    for idx, h in enumerate(header_idxs):
        header_row = rows[h]
        col_ruc = _find_col(header_row, "RUC")
        col_name = _find_col(header_row, "RAZON_SOCIAL") if _find_col(header_row, "RAZON_SOCIAL") >= 0 else _find_col(header_row, "RAZON SOCIAL")
        if min(col_ruc, col_name) < 0:
            continue
        end = header_idxs[idx + 1] if idx + 1 < len(header_idxs) else len(rows)
        for row in rows[h + 1 : end]:
            ruc = _normalize_ruc(row[col_ruc]) if col_ruc < len(row) else ""
            name = _normalize(row[col_name]) if col_name < len(row) else ""
            if ruc and name:
                name_to_ruc[name] = ruc

    # TAMANO_EMPRESA_GLOBAL
    df_tam = xl.parse("TAMANO_EMPRESA_GLOBAL").fillna("")
    if not df_tam.empty and len(df_tam.columns) >= 2:
        for _, row in df_tam.iterrows():
            ruc = _normalize_ruc(row.iloc[0])
            tam = _normalize(row.iloc[1])
            if ruc and tam:
                ruc_to_tam[ruc] = tam

    return name_to_ruc, ruc_to_tam


def _load_capacitaciones(xl: pd.ExcelFile, name_to_ruc: Dict[str, str], ruc_to_tam: Dict[str, str]):
    """
    Devuelve lista de tuples (key, data) desde la hoja CAPACITACIONES.
    key = ruc si existe, de lo contrario nombre normalizado.
    data = dict con nombre, tam, total_cap, valor, es_socio
    """
    df_raw = xl.parse("CAPACITACIONES", header=0).fillna("")
    header_row = df_raw.iloc[0].tolist()
    df = df_raw.iloc[1:].copy()
    df.columns = header_row

    headers = list(df.columns)
    pos_name = _find_col(headers, "RAZON_SOCIAL")
    pos_tam = _find_col(headers, "TAMANO")
    pos_tot = _find_col(headers, "TOTAL CAPAC")
    pos_val = _find_col(headers, "VALOR TOTAL")

    records = []
    for _, row in df.iterrows():
        # Saltar filas sin índice numérico (primera columna)
        try:
            idx_num = pd.to_numeric(row.iloc[0], errors="coerce")
        except Exception:
            idx_num = None
        if pd.isna(idx_num):
            continue

        name = str(row.iloc[pos_name]).strip() if pos_name >= 0 and pos_name < len(row) else ""
        if not name:
            continue
        norm_name = _normalize(name)
        es_socio = norm_name != "NO SOCIOS"

        ruc = name_to_ruc.get(norm_name, "")
        tam_sheet = _normalize(row.iloc[pos_tam]) if pos_tam >= 0 and pos_tam < len(row) else ""
        tam_global = ruc_to_tam.get(ruc, "") if ruc else ""
        tam = tam_sheet or tam_global

        total_cap_val = pd.to_numeric(row.iloc[pos_tot], errors="coerce") if pos_tot >= 0 and pos_tot < len(row) else 0
        valor_val = pd.to_numeric(row.iloc[pos_val], errors="coerce") if pos_val >= 0 and pos_val < len(row) else 0

        total_cap = int(total_cap_val) if pd.notna(total_cap_val) else 0
        valor = float(valor_val) if pd.notna(valor_val) else 0.0

        key = ruc or norm_name
        records.append((key, {"razon_social": name, "ruc": ruc, "tamano": tam, "total_cap": total_cap, "valor_total": valor, "es_socio": es_socio}))
    return records


def _load_capacitaciones_final(xl: pd.ExcelFile, name_to_ruc: Dict[str, str], ruc_to_tam: Dict[str, str]):
    """
    Devuelve lista de tuples (key, data) desde la hoja CAPACITACIONES_FINAL.
    Cada fila suma 1 a total_cap y su valor_pago a valor_total.
    """
    df = xl.parse("CAPACITACIONES_FINAL").fillna("")
    if df.empty:
        return []

    records = []
    for _, row in df.iterrows():
        name = str(row.get("Razon Social", "")).strip()
        if not name:
            continue
        norm_name = _normalize(name)
        es_socio = norm_name != "NO SOCIOS"

        ruc = name_to_ruc.get(norm_name, "")
        tam = ruc_to_tam.get(ruc, "")

        valor = pd.to_numeric(row.get("Valor del Pago", 0), errors="coerce")
        valor = float(valor) if pd.notna(valor) else 0.0

        key = ruc or norm_name
        records.append((key, {"razon_social": name, "ruc": ruc, "tamano": tam, "total_cap": 1, "valor_total": valor, "es_socio": es_socio}))
    return records


def _aggregate(records: List[Tuple[str, Dict]]) -> List[Dict]:
    agg: Dict[str, Dict] = {}
    for key, data in records:
        if not key:
            continue
        current = agg.get(key, {"razon_social": data.get("razon_social", ""), "ruc": data.get("ruc", ""), "tamano": "", "total_cap": 0, "valor_total": 0.0, "es_socio": data.get("es_socio", True)})
        # Prefer tamano no vacío
        if data.get("tamano"):
            current["tamano"] = data["tamano"]
        if not current.get("razon_social"):
            current["razon_social"] = data.get("razon_social", "")
        if not current.get("ruc"):
            current["ruc"] = data.get("ruc", "")
        current["total_cap"] += data.get("total_cap", 0)
        current["valor_total"] += data.get("valor_total", 0.0)
        # es_socio es AND para mantener NO SOCIOS si alguno lo marca
        current["es_socio"] = current.get("es_socio", True) and data.get("es_socio", True)
        agg[key] = current
    return list(agg.values())


def build_dataframe(base_path: str) -> pd.DataFrame:
    if not os.path.exists(base_path):
        raise FileNotFoundError(f"No se encontró el archivo base: {base_path}")

    xl = pd.ExcelFile(base_path)
    name_to_ruc, ruc_to_tam = _load_base_maps(xl)

    records = []
    records.extend(_load_capacitaciones(xl, name_to_ruc, ruc_to_tam))
    records.extend(_load_capacitaciones_final(xl, name_to_ruc, ruc_to_tam))

    aggregated = _aggregate(records)
    df = pd.DataFrame(aggregated)
    # Ordenar columnas
    cols = ["ruc", "razon_social", "tamano", "total_cap", "valor_total", "es_socio"]
    df = df[cols]
    return df


def _ensure_django():
    import django
    from django.conf import settings

    if not settings.configured:
        os.environ.setdefault("DJANGO_SETTINGS_MODULE", "capig_form.settings")
        django.setup()


def _ensure_worksheet(spreadsheet, name: str):
    from gspread.exceptions import WorksheetNotFound

    try:
        return spreadsheet.worksheet(name)
    except WorksheetNotFound:
        return spreadsheet.add_worksheet(title=name, rows=2000, cols=20)


def _write_sheet(ws, rows: List[List]):
    ws.clear()
    if rows:
        ws.update(rows)


def run(update_sheet: bool = True, sheet_name: str = "CAPACITACIONES_DASH_DATA"):
    df = build_dataframe(BASE_PATH)

    # Guardar local
    os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)
    df.to_excel(OUTPUT_PATH, index=False)
    print(f"Archivo generado: {OUTPUT_PATH}")

    if not update_sheet:
        return

    # Opcional: actualizar Google Sheets
    _ensure_django()
    from django.conf import settings
    from capig_form.services import google_sheets_service as gss

    sheet_id = os.getenv("SHEET_PATH") or getattr(settings, "SHEET_PATH", "")
    if not sheet_id:
        print("SHEET_PATH no está configurado; se omitió la actualización en Google Sheets.")
        return

    client = gss._get_client()
    ss = client.open_by_key(sheet_id)
    ws = _ensure_worksheet(ss, sheet_name)

    header = list(df.columns)
    rows = [header] + df.values.tolist()
    _write_sheet(ws, rows)
    print(f"Actualizado Google Sheet {sheet_id} -> {sheet_name}")


if __name__ == "__main__":
    run()
