import os
import re
import unicodedata
from typing import Iterable

import pandas as pd
from django.conf import settings

from capig_form.services.google_sheets_service import update_sheet_with_dataframe

DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
os.makedirs(DATA_DIR, exist_ok=True)
EXCEL_FILE = os.path.join(DATA_DIR, "datos_completos.xlsx")


def _normalize_label(value) -> str:
    text = "" if value is None else str(value)
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
    text = re.sub(r"[^0-9A-Za-z]+", "_", text)
    return text.strip("_").upper()


def _log_sheet_info(name: str, df: pd.DataFrame) -> None:
    print(f"[{name}] Filas: {len(df)}, Columnas normalizadas: {list(df.columns)}")


def _prepare_table(df: pd.DataFrame, name: str) -> pd.DataFrame:
    if df.empty:
        print(f"[{name}] Hoja vacia.")
        return pd.DataFrame()

    header = df.iloc[0].tolist()
    normalized = []
    for idx, col in enumerate(header):
        normalized_label = _normalize_label(col)
        if not normalized_label:
            normalized_label = f"COLUMN_{idx}"
        normalized.append(normalized_label)

    body = df.iloc[1:].copy()
    body.columns = normalized
    body = body.reset_index(drop=True)
    _log_sheet_info(name, body)
    return body


def _load_sheet(sheet_name: str) -> pd.DataFrame:
    try:
        raw_df = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name)
    except FileNotFoundError as exc:
        raise RuntimeError(f"No se encontro el archivo de datos: {EXCEL_FILE}") from exc
    except ValueError as exc:
        raise RuntimeError(f"La hoja '{sheet_name}' no existe en {EXCEL_FILE}") from exc

    return _prepare_table(raw_df, sheet_name)


def _clean_string(value: str) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    text = str(value).strip()
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
    return text.strip().strip("'")


def _safe_int(value) -> int:
    try:
        return int(float(value))
    except (TypeError, ValueError):
        return 0


def _safe_float(value) -> float:
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def _require_columns(df: pd.DataFrame, required: Iterable[str], sheet_name: str) -> None:
    missing = [col for col in required if col not in df.columns]
    if missing:
        raise RuntimeError(f"Columnas faltantes en hoja '{sheet_name}': {missing}")


def _extract_otro_text(row: pd.Series) -> str:
    for key in row.index:
        if key == 'OTROS':
            continue
        if 'OTRO' in key.upper():
            value = row.get(key)
            if pd.notna(value):
                cleaned = _clean_string(value)
                if cleaned:
                    return cleaned
    return ''


def _is_valid_company(name: str) -> bool:
    if not name:
        return False
    text = name.strip()
    return bool(text) and text.upper() != 'NO SOCIOS'


def _is_positive_flag(value) -> bool:
    if pd.isna(value):
        return False
    if isinstance(value, (int, float)):
        return value > 0
    text = str(value).strip()
    if not text:
        return False
    if text.upper() in ('X', 'SI', 'YES', 'TRUE', '1'):
        return True
    try:
        return float(text) > 0
    except ValueError:
        return False


def generar_diagnostico_y_subir():
    base_datos = _load_sheet("BASE DE DATOS")
    diagnosticos = _load_sheet("DIAGNOSTICOS")
    legal1 = _load_sheet("LEGAL")
    legal2 = _load_sheet("LEGAL (2)")
    if 'INTELECTUAL' in legal2.columns and 'PROPIEDAD_INTELECTUAL' not in legal2.columns:
        print("Normalizando columna INTELECTUAL -> PROPIEDAD_INTELECTUAL en LEGAL (2)")
        legal2 = legal2.rename(columns={'INTELECTUAL': 'PROPIEDAD_INTELECTUAL'})
    capacitaciones = _load_sheet("CAPACITACIONES")

    def _collect_empresas():
        """Recolecta todas las empresas de Diagnosticos y Base de Datos para no omitir ninguna."""
        empresas = set()
        for raw in diagnosticos.get('RAZON_SOCIAL', []):
            nombre = obtener_nombre(raw)
            if _is_valid_company(nombre):
                empresas.add(nombre)
        for raw in base_datos.get('RAZON_SOCIAL', []):
            nombre = obtener_nombre(raw)
            if _is_valid_company(nombre):
                empresas.add(nombre)
        return empresas

    tipo_map = {
        'LEAN': 'lean',
        'ESTRATEGIA': 'estrategia',
        'LEGAL': 'legal',
        'AMBIENTE': 'ambiente',
        'RRHH': 'rrhh'
    }

    required_diag_cols = ['RAZON_SOCIAL', 'DIAGNOSTICO'] + list(tipo_map.keys())
    _require_columns(base_datos, ['RUC', 'RAZON_SOCIAL'], 'BASE DE DATOS')
    _require_columns(diagnosticos, required_diag_cols, 'DIAGNOSTICOS')

    if base_datos.empty:
        raise RuntimeError("La hoja 'BASE DE DATOS' no contiene informacion.")

    if diagnosticos.empty:
        raise RuntimeError("La hoja 'DIAGNOSTICOS' no contiene informacion.")

    ruc_to_nombre = {}
    if 'RUC' in base_datos.columns and 'RAZON_SOCIAL' in base_datos.columns:
        base_datos['RUC'] = base_datos['RUC'].apply(_clean_string)
        base_datos['RAZON_SOCIAL'] = base_datos['RAZON_SOCIAL'].apply(_clean_string)
        ruc_to_nombre = dict(zip(base_datos['RUC'], base_datos['RAZON_SOCIAL']))

    def obtener_nombre(valor):
        cleaned = _clean_string(valor)
        return ruc_to_nombre.get(cleaned, cleaned)

    rows = []
    for _, row in diagnosticos.iterrows():
        entidad = obtener_nombre(row.get('RAZON_SOCIAL', ''))
        if not _is_valid_company(entidad):
            continue
        for col, tipo in tipo_map.items():
            flag = row.get(col, '')
            if not _is_positive_flag(flag):
                continue
            rows.append({
                'Razon Social': entidad,
                'Tipo de Diagnostico': tipo,
                'Subtipo de Diagnostico': 'PENDIENTE' if tipo == 'legal' else '',
                'Otros Subtipo': '',
                'Se Diagnostico': 'Si',
                'Fecha': '',
                'Hora': ''
            })

    LEGAL_SUBTYPES = {
        'LABORAL': 'laboral',
        'SOCIETARIO': 'societario',
        'PROPIEDAD_INTELECTUAL': 'propiedad intelectual',
        'OTROS': 'otros',
        'CONTACTO': 'contacto',
    }

    def procesar_legal(df_legal, sheet_label: str):
        if df_legal.empty:
            return

        # Solo exigir columnas que realmente existen en esta hoja
        columnas_en_hoja = df_legal.columns.tolist()
        columnas_requeridas = ['RAZON_SOCIAL'] + [col for col in LEGAL_SUBTYPES if col in columnas_en_hoja]
        _require_columns(df_legal, columnas_requeridas, sheet_label)

        for _, row in df_legal.iterrows():
            razon_social = obtener_nombre(row.get('RAZON_SOCIAL', ''))
            if not _is_valid_company(razon_social):
                continue
            for column, label in LEGAL_SUBTYPES.items():
                if column not in columnas_en_hoja:
                    continue
                valor = row.get(column, '')
                if not _is_positive_flag(valor):
                    continue
                se_diag = 'Si'
                otros = ''
                if column == 'OTROS':
                    otros = _extract_otro_text(row)

                rows.append({
                    'Razon Social': razon_social,
                    'Tipo de Diagnostico': 'legal',
                    'Subtipo de Diagnostico': label,
                    'Otros Subtipo': otros,
                    'Se Diagnostico': se_diag,
                    'Fecha': '',
                    'Hora': ''
                })

    procesar_legal(legal1, 'LEGAL')
    procesar_legal(legal2, 'LEGAL (2)')

    df_final = pd.DataFrame(rows).drop_duplicates().fillna("")

    # Incluir empresas sin diagnosticos positivos con un marcador explicito.
    empresas_totales = _collect_empresas()
    empresas_en_df = set(df_final['Razon Social']) if not df_final.empty else set()
    pendientes_empresas = empresas_totales - empresas_en_df
    if pendientes_empresas:
        print(f"Agregando {len(pendientes_empresas)} empresas sin diagnosticos positivos con Tipo=ninguno")
        adicionales = pd.DataFrame([{
            'Razon Social': empresa,
            'Tipo de Diagnostico': 'ninguno',
            'Subtipo de Diagnostico': '',
            'Otros Subtipo': '',
            'Se Diagnostico': 'No',
            'Fecha': '',
            'Hora': ''
        } for empresa in sorted(pendientes_empresas)])
        df_final = pd.concat([df_final, adicionales], ignore_index=True)

    print(f"Filas procesadas antes de limpieza: {len(rows)}")
    print(f"Filas finales para subir: {len(df_final)}")
    if not df_final.empty:
        counts_por_tipo = df_final['Tipo de Diagnostico'].value_counts().to_dict()
        print(f"Totales por tipo de diagnostico: {counts_por_tipo}")
    print("Filas totales DIAGNOSTICO_FINAL:", len(df_final))
    if not df_final.empty:
        print("Primeras 20 filas de DIAGNOSTICO_FINAL:")
        print(df_final.head(20))
    else:
        print("DIAGNOSTICO_FINAL esta vacio, revisa los datos de origen.")

    df_capacitaciones_final = _build_capacitaciones(capacitaciones)

    sheet_id = settings.SHEET_PATH
    update_sheet_with_dataframe(sheet_id, "DIAGNOSTICO_FINAL", df_final)
    update_sheet_with_dataframe(sheet_id, "CAPACITACIONES_FINAL", df_capacitaciones_final)


def _build_capacitaciones(capacitaciones: pd.DataFrame) -> pd.DataFrame:
    columns = [
        'Razon Social',
        'Nombre de la Capacitacion',
        'Tipo de Capacitacion',
        'Valor del Pago',
        'Fecha',
        'Hora',
    ]

    df_cap = pd.DataFrame(columns=columns)
    print(f"Filas de capacitaciones procesadas: {len(df_cap)}")
    if not df_cap.empty:
        print(df_cap.head())
    return df_cap
