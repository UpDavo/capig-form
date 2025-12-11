import os
import re
from collections import defaultdict
from typing import Dict, List, Tuple

import pandas as pd


def _ensure_django():
    import django
    from django.conf import settings

    if not settings.configured:
        os.environ.setdefault("DJANGO_SETTINGS_MODULE", "capig_form.settings")
        django.setup()


def _normalize(text: str) -> str:
    txt = str(text or "").strip().upper()
    txt = (
        txt.replace("Á", "A")
        .replace("É", "E")
        .replace("Í", "I")
        .replace("Ó", "O")
        .replace("Ú", "U")
        .replace("Ñ", "N")
    )
    txt = re.sub(r"\s+", "_", txt)
    return txt


def _find_col(headers: List[str], needle: str) -> int:
    target = _normalize(needle)
    for idx, h in enumerate(headers):
        if target in _normalize(h):
            return idx
    return -1


def _to_float(val) -> float:
    if val is None:
        return 0.0
    if isinstance(val, (int, float)):
        try:
            return float(val)
        except Exception:
            return 0.0
    txt = str(val).strip()
    if not txt:
        return 0.0
    try:
        return float(txt.replace(",", "").replace("$", "").replace(" ", ""))
    except Exception:
        try:
            return float(re.sub(r"[^\d\.-]", "", txt))
        except Exception:
            return 0.0


def _sum_year_columns(headers: List[str], row: List) -> float:
    total = 0.0
    for idx, h in enumerate(headers):
        norm = _normalize(h)
        # Solo años puros de 4 dígitos (2019, 2020, 2021, 2022, 2023...), ignora columnas con prefijo T
        if re.fullmatch(r"\d{4}", norm):
            if idx < len(row):
                total += _to_float(row[idx])
    return total


def _collect_sector_map(ss) -> Dict[str, str]:
    sector_map: Dict[str, str] = {}
    try:
        sh = ss.worksheet("SECTOR")
    except Exception:
        return sector_map

    values = sh.get_all_values()
    if not values:
        return sector_map
    headers = values[0]
    col_ruc = _find_col(headers, "RUC")
    col_sec = _find_col(headers, "SECTOR")
    if min(col_ruc, col_sec) < 0:
        return sector_map
    for row in values[1:]:
        if col_ruc >= len(row) or col_sec >= len(row):
            continue
        ruc = str(row[col_ruc]).replace("'", "").replace('"', "").strip()
        sec = str(row[col_sec]).strip()
        if ruc:
            sector_map[ruc] = sec
    return sector_map


def _collect_base_data(data_bd: List[List], sector_map: Dict[str, str]) -> Dict[str, Dict]:
    result: Dict[str, Dict] = {}

    def process_block(header_row: List[str], rows: List[List[str]]):
        col_ruc = _find_col(header_row, "RUC")
        col_emp = _find_col(header_row, "RAZON_SOCIAL") if _find_col(header_row, "RAZON_SOCIAL") >= 0 else _find_col(header_row, "RAZON SOCIAL")
        col_tam = _find_col(header_row, "TAMANO")
        col_estado = _find_col(header_row, "ESTADO")
        col_colab = _find_col(header_row, "COLABORADORES") if _find_col(header_row, "COLABORADORES") >= 0 else _find_col(header_row, "NO_COLABORADORES")
        col_semaforo = _find_col(header_row, "SEMAFORO")
        col_sector = _find_col(header_row, "SECTOR")

        for row in rows:
            if not row or all(not c for c in row):
                continue
            ruc = str(row[col_ruc] if col_ruc >= 0 and col_ruc < len(row) else "").replace("'", "").replace('"', "").strip()
            if not ruc:
                continue
            empresa = str(row[col_emp]) if col_emp >= 0 and col_emp < len(row) else ""
            tamano = str(row[col_tam]) if col_tam >= 0 and col_tam < len(row) else ""
            estado = str(row[col_estado]) if col_estado >= 0 and col_estado < len(row) else ""
            colab = row[col_colab] if col_colab >= 0 and col_colab < len(row) else ""
            semaforo = str(row[col_semaforo]) if col_semaforo >= 0 and col_semaforo < len(row) else ""
            sector_cell = str(row[col_sector]) if col_sector >= 0 and col_sector < len(row) else ""
            ventas_hist = _sum_year_columns(header_row, row)

            sec = sector_map.get(ruc, sector_cell)

            current = result.get(ruc, {})
            # Prefer non-empty values
            def pick(new, old):
                return new if str(new).strip() else old

            result[ruc] = {
                "empresa": pick(empresa, current.get("empresa", "")),
                "tamano": pick(tamano, current.get("tamano", "")),
                "estado": pick(estado, current.get("estado", "")),
                "colab": pick(colab, current.get("colab", "")),
                "semaforo": pick(semaforo, current.get("semaforo", "")),
                "sector": pick(sec, current.get("sector", "")),
                "ventas_hist": current.get("ventas_hist", 0.0) + ventas_hist,
            }

    # Header rows: row 2 (idx1) and row 314 (idx313) per spec, but detect flex
    header_idxs = []
    for i, row in enumerate(data_bd):
        if not row:
            continue
        if _find_col(row, "RUC") >= 0 and _find_col(row, "TAMANO") >= 0:
            header_idxs.append(i)
    h1 = 1 if 1 in header_idxs else (header_idxs[0] if header_idxs else 1)
    h2 = 313 if 313 in header_idxs else (header_idxs[1] if len(header_idxs) > 1 else None)

    if h1 is not None and h1 < len(data_bd):
        process_block(data_bd[h1], data_bd[h1 + 1 : 312])
    if h2 is not None and h2 < len(data_bd):
        process_block(data_bd[h2], data_bd[h2 + 1 :])

    return result


def _collect_ventas_nuevas(data_ventas: List[List]) -> Dict[str, float]:
    if not data_ventas:
        return {}
    headers = data_ventas[0]
    col_ruc = _find_col(headers, "RUC")
    col_monto = _find_col(headers, "MONTO_ESTIMADO")
    if col_monto < 0:
        col_monto = _find_col(headers, "MONTO")
    if min(col_ruc, col_monto) < 0:
        return {}
    agg: Dict[str, float] = defaultdict(float)
    for row in data_ventas[1:]:
        if col_ruc >= len(row) or col_monto >= len(row):
            continue
        ruc = str(row[col_ruc]).replace("'", "").replace('"', "").strip()
        if not ruc:
            continue
        monto = _to_float(row[col_monto])
        agg[ruc] += monto
    return agg


def _ensure_worksheet(ss, name: str):
    from gspread.exceptions import WorksheetNotFound

    try:
        return ss.worksheet(name)
    except WorksheetNotFound:
        return ss.add_worksheet(title=name, rows=2000, cols=20)


def _write_sheet(ws, rows: List[List]):
    ws.clear()
    if rows:
        ws.update(rows)


def run():
    _ensure_django()
    from django.conf import settings
    from capig_form.services import google_sheets_service as gss

    sheet_id = os.getenv("SHEET_PATH") or getattr(settings, "SHEET_PATH", "")
    if not sheet_id:
        raise RuntimeError("SHEET_PATH no esta configurado.")

    client = gss._get_client()
    ss = client.open_by_key(sheet_id)

    hoja_bd = ss.worksheet("BASE DE DATOS")
    hoja_ventas = ss.worksheet("VENTAS_AFILIADOS")

    data_bd = hoja_bd.get_all_values()
    data_ventas = hoja_ventas.get_all_values()

    sector_map = _collect_sector_map(ss)
    base_data = _collect_base_data(data_bd, sector_map)
    ventas_nuevas = _collect_ventas_nuevas(data_ventas)

    rows = [["RUC", "Empresa", "Tamano", "Estado", "Sector", "Colaboradores", "Semaforo", "Ventas Totales"]]
    for ruc, info in base_data.items():
        ventas_hist = info.get("ventas_hist", 0.0)
        ventas_new = ventas_nuevas.get(ruc, 0.0)
        ventas_total = ventas_hist + ventas_new
        rows.append(
            [
                ruc,
                info.get("empresa", ""),
                info.get("tamano", ""),
                info.get("estado", ""),
                info.get("sector", ""),
                info.get("colab", ""),
                info.get("semaforo", ""),
                ventas_total,
            ]
        )

    hoja_dash = _ensure_worksheet(ss, "DASH_DATA")
    _write_sheet(hoja_dash, rows)


if __name__ == "__main__":
    run()
