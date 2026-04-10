import logging
import os
import re
from datetime import datetime, timedelta
from typing import Dict

from django.conf import settings
from gspread.utils import rowcol_to_a1

from capig_form.services.google_sheets_service import (
    ensure_row_capacity,
    find_first_empty_row,
    get_google_sheet,
)

EXPECTED_BASE_HEADERS = [
    "RUC",
    "RAZON_SOCIAL",
    "CIUDAD",
    "FECHA_AFILIACION",
]


def limpiar_ruc(valor):
    """Normaliza el RUC dejando solo digitos."""
    txt = (
        str(valor)
        .replace("'", "")
        .replace('"', "")
        .replace("\u00a0", " ")
        .strip()
    )
    return re.sub(r"\D", "", txt)


def _ruc_compare_key(valor):
    """
    Normaliza RUC para comparaciones robustas.
    Ignora ceros a la izquierda para tolerar conversiones de Sheets.
    """
    ruc = limpiar_ruc(valor)
    if re.fullmatch(r"\d+", ruc or ""):
        return ruc.lstrip("0") or "0"
    return ""


def _normalize_row_keys(row: dict):
    """Devuelve un diccionario con llaves normalizadas."""
    return {(k or "").strip().upper(): v for k, v in (row or {}).items()}


def _normalize_header_key(value):
    """Normaliza encabezados para comparaciones."""
    return str(value or "").replace("\u00a0", " ").strip().upper()


def excel_serial_to_iso(valor):
    """
    Convierte serial de Excel (float/int o str numerico) a YYYY-MM-DD.
    Si no aplica, devuelve la cadena limpia.
    """
    if isinstance(valor, str):
        val_strip = valor.strip()
        if re.fullmatch(r"-?\d+(\.\d+)?", val_strip):
            try:
                valor = float(val_strip)
            except Exception:
                return val_strip
        else:
            return val_strip

    if isinstance(valor, (int, float)) and valor:
        try:
            base = datetime(1899, 12, 30)
            return (base + timedelta(days=float(valor))).date().isoformat()
        except Exception:
            return str(valor)
    return str(valor).strip() if valor is not None else ""


def _get_estado_sheet():
    """Obtiene la hoja ESTADO_SOCIO."""
    sheet_id = os.getenv("SHEET_PATH") or getattr(settings, "SHEET_PATH", "")
    if not sheet_id:
        raise RuntimeError("SHEET_PATH no esta configurado.")
    return get_google_sheet(sheet_id, "ESTADO_SOCIO")


def _get_base_datos_sheet():
    """Obtiene la hoja SOCIOS."""
    sheet_id = os.getenv("SHEET_PATH") or getattr(settings, "SHEET_PATH", "")
    if not sheet_id:
        raise RuntimeError("SHEET_PATH no esta configurado.")
    return get_google_sheet(sheet_id, "SOCIOS")


def _get_all_records_flexible(sheet, head=2, required_keys=("RUC",)):
    """
    Lee registros probando encabezados en orden: head solicitado, 1 y 2.
    Valida que la cabecera se parezca a la esperada.
    """
    required = {_normalize_header_key(k) for k in (required_keys or []) if k}

    candidate_heads = []
    for value in (head, 1, 2):
        if value and value not in candidate_heads:
            candidate_heads.append(value)

    for candidate in candidate_heads:
        if required:
            try:
                header_values = sheet.row_values(candidate)
            except Exception:
                continue
            header_keys = {
                _normalize_header_key(v) for v in header_values if str(v or "").strip()
            }
            if header_keys and not (header_keys & required):
                continue

        try:
            return sheet.get_all_records(
                head=candidate,
                value_render_option="UNFORMATTED_VALUE",
                numericise_ignore=["all"],
            )
        except Exception:
            continue
    return []


def buscar_afiliado_por_ruc(ruc):
    """Busca primero en ESTADO_SOCIO; si falta info, completa desde SOCIOS."""
    ruc_key = _ruc_compare_key(ruc)

    estado_sheet = _get_estado_sheet()
    estado_rows = _get_all_records_flexible(estado_sheet, head=1)
    afiliado = next(
        (
            row
            for row in estado_rows
            if _ruc_compare_key(row.get("RUC", "")) == ruc_key
        ),
        None,
    )

    if afiliado and all(
        afiliado.get(key) for key in ["RAZON_SOCIAL", "CIUDAD", "FECHA_AFILIACION", "ESTADO"]
    ):
        return {
            "razon_social": afiliado.get("RAZON_SOCIAL", ""),
            "ciudad": afiliado.get("CIUDAD", ""),
            "fecha_afiliacion": excel_serial_to_iso(afiliado.get("FECHA_AFILIACION", "")),
            "estado": afiliado.get("ESTADO", ""),
        }

    base_sheet = _get_base_datos_sheet()
    base_rows = _get_all_records_flexible(base_sheet, head=2)
    base_row = next(
        (
            row
            for row in base_rows
            if _ruc_compare_key(row.get("RUC", "")) == ruc_key
        ),
        None,
    )

    if afiliado:
        return {
            "razon_social": afiliado.get("RAZON_SOCIAL", ""),
            "ciudad": afiliado.get("CIUDAD", ""),
            "fecha_afiliacion": excel_serial_to_iso(afiliado.get("FECHA_AFILIACION", "")),
            "estado": afiliado.get("ESTADO", ""),
        }

    if not base_row:
        return None

    return {
        "razon_social": base_row.get("RAZON_SOCIAL", ""),
        "ciudad": base_row.get("CIUDAD", ""),
        "fecha_afiliacion": excel_serial_to_iso(base_row.get("FECHA_AFILIACION", "")),
        "estado": "",
    }


def actualizar_estado_afiliado(ruc, nuevo_estado):
    """Actualiza el estado del afiliado y crea fila si no existe."""
    sheet = _get_estado_sheet()
    data = sheet.get_all_records(
        value_render_option="UNFORMATTED_VALUE",
        numericise_ignore=["all"],
    )
    header = [col.strip().upper() for col in sheet.row_values(1)]

    def _col_index(nombre):
        try:
            return header.index(nombre) + 1
        except ValueError:
            return None

    col_estado = _col_index("ESTADO")
    col_actualizacion = _col_index("ACTUALIZACION_ESTADO")
    encontrado = False

    for idx, row in enumerate(data, start=2):
        if _ruc_compare_key(row.get("RUC", "")) == _ruc_compare_key(ruc):
            encontrado = True
            if col_estado:
                sheet.update_cell(idx, col_estado, nuevo_estado)
            if col_actualizacion:
                sheet.update_cell(
                    idx,
                    col_actualizacion,
                    datetime.now().strftime("%Y-%m-%d %H:%M"),
                )
            break

    if not encontrado:
        base_sheet = _get_base_datos_sheet()
        base_rows = _get_all_records_flexible(base_sheet, head=2)
        base_row = next(
            (
                row
                for row in base_rows
                if _ruc_compare_key(row.get("RUC", "")) == _ruc_compare_key(ruc)
            ),
            {},
        )
        new_row = [
            limpiar_ruc(ruc),
            base_row.get("RAZON_SOCIAL", ""),
            excel_serial_to_iso(base_row.get("FECHA_AFILIACION", "")),
            nuevo_estado,
            base_row.get("CIUDAD", ""),
            datetime.now().strftime("%Y-%m-%d %H:%M"),
        ]

        header_len = max(len(header), len(new_row))
        target_row = find_first_empty_row(sheet, start_row=2)
        ensure_row_capacity(sheet, target_row)

        if len(new_row) < header_len:
            new_row += [""] * (header_len - len(new_row))
        elif len(new_row) > header_len:
            new_row = new_row[:header_len]

        start_cell = rowcol_to_a1(target_row, 1)
        end_cell = rowcol_to_a1(target_row, header_len)
        sheet.update(f"{start_cell}:{end_cell}", [new_row], value_input_option="USER_ENTERED")


def buscar_afiliado_por_ruc_base_datos(ruc):
    """Busca un afiliado unicamente en la hoja SOCIOS."""
    ruc_key = _ruc_compare_key(ruc)
    sheet = _get_base_datos_sheet()
    rows = _get_all_records_flexible(sheet, head=2)
    for row in rows:
        if _ruc_compare_key(row.get("RUC", "")) == ruc_key:
            return {
                "razon_social": row.get("RAZON_SOCIAL", ""),
                "ciudad": row.get("CIUDAD", ""),
                "fecha_afiliacion": excel_serial_to_iso(row.get("FECHA_AFILIACION", "")),
            }
    return None


def listar_empresas_socias():
    """
    Devuelve la lista de empresas registradas en SOCIOS.

    Se usa para poblar selects de formularios operativos sin depender
    del indice de la hoja ni de columnas fijas.
    """
    try:
        sheet = _get_base_datos_sheet()
        rows = _get_all_records_flexible(
            sheet,
            head=2,
            required_keys=("RUC", "RAZON_SOCIAL"),
        )
    except Exception:
        logging.exception("No se pudo cargar la lista de empresas desde SOCIOS.")
        return []

    empresas = {}
    for row in rows:
        row_norm = _normalize_row_keys(row)
        razon_social = str(row_norm.get("RAZON_SOCIAL", "") or "").strip()
        ruc = limpiar_ruc(row_norm.get("RUC", ""))

        if not razon_social:
            continue

        dedupe_key = _ruc_compare_key(ruc) or razon_social.casefold()
        current = empresas.get(dedupe_key)
        if not current:
            empresas[dedupe_key] = {
                "razon_social": razon_social,
                "ruc": ruc,
            }
            continue

        if not current.get("ruc") and ruc:
            current["ruc"] = ruc

    return sorted(
        empresas.values(),
        key=lambda item: item.get("razon_social", "").casefold(),
    )


def obtener_ventas_por_ruc(ruc):
    """Obtiene ventas historicas del afiliado desde VENTAS_SOCIO y, si no hay, desde SOCIOS."""
    ruc_key = _ruc_compare_key(ruc)
    if not ruc_key:
        return []

    sheet_id = os.getenv("SHEET_PATH") or getattr(settings, "SHEET_PATH", "")
    if not sheet_id:
        return []

    try:
        ventas_sheet = get_google_sheet(sheet_id, "VENTAS_SOCIO")
    except Exception:
        ventas_sheet = None

    rows = []
    if ventas_sheet is not None:
        rows = _get_all_records_flexible(
            ventas_sheet,
            head=1,
            required_keys=("RUC", "RAZON_SOCIAL", "ANIO", "AÑO", "ANO"),
        )

    ventas = []
    for row in rows:
        row_norm = _normalize_row_keys(row)
        if _ruc_compare_key(row_norm.get("RUC", "")) != ruc_key:
            continue

        anio = str(
            row_norm.get("ANIO") or row_norm.get("AÑO") or row_norm.get("ANO") or ""
        ).strip()
        comparativo = row_norm.get("COMPARATIVO", "")
        bruto_ventas = (
            row_norm.get("VENTAS_ESTIMADAS")
            or row_norm.get("MONTO_ESTIMADO")
            or row_norm.get("MONTO_VENTAS")
            or row.get("MONTO_VENTAS ")
            or row_norm.get("VENTAS_ESTIMADA")
            or ""
        )
        if isinstance(bruto_ventas, (int, float)):
            ventas_estimadas = str(bruto_ventas)
        else:
            ventas_estimadas = (bruto_ventas or "").strip()

        fecha_registro = excel_serial_to_iso(
            row_norm.get("FECHA_REGISTRO", "") or row_norm.get("FECHA", "")
        )
        ventas.append(
            {
                "anio": anio,
                "comparativo": comparativo,
                "ventas_estimadas": ventas_estimadas,
                "fecha_registro": fecha_registro,
            }
        )

    try:
        base_sheet = get_google_sheet(sheet_id, "SOCIOS")
        base_rows = _get_all_records_flexible(base_sheet, head=2)
    except Exception:
        base_rows = []

    if base_rows:
        try:
            base_row = next(
                (
                    row
                    for row in base_rows
                    if _ruc_compare_key(_normalize_row_keys(row).get("RUC", "")) == ruc_key
                ),
                None,
            )
        except Exception:
            base_row = None

        if base_row:
            base_row_norm = _normalize_row_keys(base_row)
            existing_years = {v.get("anio") for v in ventas if v.get("anio")}
            for key, value in base_row_norm.items():
                key_str = (key or "").replace("\u00a0", "").strip()
                if not key_str or not re.fullmatch(r"\d{4}", key_str):
                    continue
                if key_str in existing_years:
                    continue
                if isinstance(value, (int, float)):
                    val_str = str(value)
                elif isinstance(value, str):
                    val_str = value.strip()
                else:
                    val_str = ""
                if val_str in ("", None):
                    continue
                ventas.append(
                    {
                        "anio": key_str,
                        "comparativo": "",
                        "ventas_estimadas": val_str,
                        "fecha_registro": "",
                    }
                )

    ventas.sort(key=lambda value: value.get("anio") or "", reverse=True)
    return ventas


def guardar_ventas_afiliado(data: Dict[str, str]):
    """
    Inserta un registro en la hoja VENTAS_SOCIO con el orden esperado.
    """
    logging.info("Datos recibidos para guardar ventas: %s", data)

    sheet_id = os.getenv("SHEET_PATH") or getattr(settings, "SHEET_PATH", "")
    if not sheet_id:
        raise RuntimeError("SHEET_PATH no esta configurado.")

    sheet = get_google_sheet(sheet_id, "VENTAS_SOCIO")
    ruc_norm = limpiar_ruc(data.get("ruc", ""))
    ruc_text = f"'{ruc_norm}" if re.fullmatch(r"\d+", ruc_norm or "") else ruc_norm

    fila = [
        ruc_text,
        data.get("razon_social", ""),
        data.get("ciudad", ""),
        data.get("fecha_afiliacion", ""),
        data.get("registro_ventas", ""),
        data.get("comparativo", ""),
        data.get("ventas_estimadas", ""),
        data.get("observaciones", ""),
        datetime.now().strftime("%Y-%m-%d %H:%M"),
        data.get("anio", ""),
    ]

    if len(fila) != 10:
        raise ValueError(f"Fila con columnas inesperadas: {fila}")

    next_row = find_first_empty_row(sheet, start_row=2)
    ensure_row_capacity(sheet, next_row)
    start = f"A{next_row}"
    end = f"J{next_row}"
    sheet.update(f"{start}:{end}", [fila], value_input_option="USER_ENTERED")
    try:
        sheet.format(
            f"D2:D{next_row}",
            {"numberFormat": {"type": "DATE", "pattern": "dd/MM/yyyy"}},
        )
    except Exception:
        logging.warning("No se pudo aplicar formato de fecha a la columna D en VENTAS_SOCIO.")
