"""
Genera la tabla de gerentes por género y tamaño y la escribe en Google Sheets.

No modifica otras lógicas ni toca más archivos; se apoya en el job local
`genero_gerentes_job` para leer y procesar el Excel.
"""
import os
from typing import List

import pandas as pd

from capig_form.services.genero_gerentes_job import BASE_PATH, TAMANO_ORDER, build_dataframe, summarize


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


def _df_to_rows(resumen: pd.DataFrame) -> List[List]:
    header = ["Tamano", "Femenino", "Masculino", "Total", "% Femenino", "% Masculino"]
    rows = [header]

    def _append_row(tam: str):
        row = resumen[resumen["tamano"] == tam]
        if row.empty:
            return
        r = row.iloc[0]
        rows.append(
            [
                tam,
                int(r["femenino"]),
                int(r["masculino"]),
                int(r["total"]),
                float(r["pct_femenino"]),
                float(r["pct_masculino"]),
            ]
        )

    for tam in TAMANO_ORDER:
        _append_row(tam)
    _append_row("GLOBAL")
    return rows


def _crosstab_rows(df: pd.DataFrame) -> List[List]:
    ct = df.groupby(["TAM_G", "genero_norm"]).size().unstack(fill_value=0)
    header = ["Tamano"] + list(ct.columns)
    rows = [header]
    for tam in TAMANO_ORDER:
        if tam not in ct.index:
            continue
        row = [tam]
        for col in ct.columns:
            row.append(int(ct.loc[tam, col]))
        rows.append(row)
    rows.append(["GLOBAL"] + [int(ct[col].sum()) for col in ct.columns])
    return rows


def _write_sheet(ws, rows: List[List]):
    ws.clear()
    if rows:
        ws.update(rows)


def run(sheet_name: str = "GERENTES_GENERO", crosstab_sheet: str = "GERENTES_GENERO_CROSSTAB"):
    _ensure_django()
    from django.conf import settings
    from capig_form.services import google_sheets_service as gss

    sheet_id = os.getenv("SHEET_PATH") or getattr(settings, "SHEET_PATH", "")
    if not sheet_id:
        raise RuntimeError("SHEET_PATH no esta configurado.")
    if not os.path.exists(BASE_PATH):
        raise FileNotFoundError(f"No se encontró el Excel base: {BASE_PATH}")

    df = build_dataframe(BASE_PATH)
    resumen = summarize(df)

    client = gss._get_client()
    ss = client.open_by_key(sheet_id)

    ws_resumen = _ensure_worksheet(ss, sheet_name)
    ws_crosstab = _ensure_worksheet(ss, crosstab_sheet)

    _write_sheet(ws_resumen, _df_to_rows(resumen))
    _write_sheet(ws_crosstab, _crosstab_rows(df))

    print(f"Actualizado Google Sheet {sheet_id} -> {sheet_name} y {crosstab_sheet}")


if __name__ == "__main__":
    run()
