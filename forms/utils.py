import os
import logging
import re
from datetime import datetime
from typing import Dict

from django.conf import settings
from gspread.utils import rowcol_to_a1

from capig_form.services.google_sheets_service import (
    get_google_sheet,
    find_first_empty_row,
)

# Encabezados mínimos usados en la hoja SOCIOS
EXPECTED_BASE_HEADERS = [
    "RUC",
    "RAZON_SOCIAL",
    "CIUDAD",
    "FECHA_AFILIACION",
]


def limpiar_ruc(valor):
    """Normaliza el RUC removiendo comillas, espacios y NBSP."""
    return (
        str(valor)
        .replace("'", "")
        .replace('"', "")
        .replace("\u00a0", " ")
        .strip()
    )


def _get_estado_sheet():
    """Obtiene la hoja de ESTADO_SOCIO desde Google Sheets."""
    sheet_id = os.getenv("SHEET_PATH") or getattr(settings, "SHEET_PATH", "")
    if not sheet_id:
        raise RuntimeError("SHEET_PATH no esta configurado.")
    return get_google_sheet(sheet_id, "ESTADO_SOCIO")


def _get_base_datos_sheet():
    """Obtiene la hoja SOCIOS desde Google Sheets."""
    sheet_id = os.getenv("SHEET_PATH") or getattr(settings, "SHEET_PATH", "")
    if not sheet_id:
        raise RuntimeError("SHEET_PATH no esta configurado.")
    return get_google_sheet(sheet_id, "SOCIOS")


def _get_all_records_flexible(sheet, head=2):
    """
    Lee registros intentando con el head indicado y, si falla (sheet vacío
    o encabezados cambiados), vuelve a head=1. Devuelve lista vacía si nada funciona.
    """
    for h in (head, 1):
        try:
            return sheet.get_all_records(
                head=h,
                value_render_option="UNFORMATTED_VALUE",
                numericise_ignore=["all"],
            )
        except Exception:
            continue
    return []


def buscar_afiliado_por_ruc(ruc):
    """
    Busca primero en ESTADO_SOCIO; si falta info, completa desde SOCIOS.
    """
    ruc = limpiar_ruc(ruc)

    estado_sheet = _get_estado_sheet()
    estado_rows = _get_all_records_flexible(estado_sheet, head=1)

    afiliado = next(
        (row for row in estado_rows if limpiar_ruc(row.get("RUC", "")) == ruc),
        None,
    )

    if afiliado and all(
        afiliado.get(key) for key in ["RAZON_SOCIAL", "CIUDAD", "FECHA_AFILIACION", "ESTADO"]
    ):
        return {
            "razon_social": afiliado.get("RAZON_SOCIAL", ""),
            "ciudad": afiliado.get("CIUDAD", ""),
            "fecha_afiliacion": afiliado.get("FECHA_AFILIACION", ""),
            "estado": afiliado.get("ESTADO", ""),
        }

    base_sheet = _get_base_datos_sheet()
    base_rows = _get_all_records_flexible(base_sheet, head=2)

    base_row = next(
        (row for row in base_rows if limpiar_ruc(row.get("RUC", "")) == ruc),
        None,
    )

    if afiliado:
        return {
            "razon_social": afiliado.get("RAZON_SOCIAL", ""),
            "ciudad": afiliado.get("CIUDAD", ""),
            "fecha_afiliacion": afiliado.get("FECHA_AFILIACION", ""),
            "estado": afiliado.get("ESTADO", ""),
        }

    if not base_row:
        return None

    return {
        "razon_social": base_row.get("RAZON_SOCIAL", ""),
        "ciudad": base_row.get("CIUDAD", ""),
        "fecha_afiliacion": base_row.get("FECHA_AFILIACION", ""),
        "estado": "",
    }


def actualizar_estado_afiliado(ruc, nuevo_estado):
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
        if limpiar_ruc(row.get("RUC", "")) == limpiar_ruc(ruc):
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

    # Si no se encontro el RUC, agregar nueva fila con datos base y estado actualizado
    if not encontrado:
        base_sheet = _get_base_datos_sheet()
        base_rows = _get_all_records_flexible(base_sheet, head=2)
        base_row = next(
            (row for row in base_rows if limpiar_ruc(
                row.get("RUC", "")) == limpiar_ruc(ruc)),
            {},
        )
        # Orden esperado: RUC | RAZON_SOCIAL | FECHA_AFILIACION | ESTADO | CIUDAD | ACTUALIZACION_ESTADO
        new_row = [
            limpiar_ruc(ruc),
            base_row.get("RAZON_SOCIAL", ""),
            base_row.get("FECHA_AFILIACION", ""),
            nuevo_estado,
            base_row.get("CIUDAD", ""),
            datetime.now().strftime("%Y-%m-%d %H:%M"),
        ]

        header_len = max(len(header), len(new_row))
        # Buscar la primera fila realmente vacía (sin celdas con texto)
        values = sheet.get_all_values(value_render_option="UNFORMATTED_VALUE")
        target_row = None
        for idx, row in enumerate(values[1:], start=2):
            if not any((cell or "").strip() for cell in row):
                target_row = idx
                break
        if target_row is None:
            target_row = len(values) + 1

        # Ajustar tamaño al header
        if len(new_row) < header_len:
            new_row += [""] * (header_len - len(new_row))
        elif len(new_row) > header_len:
            new_row = new_row[:header_len]

        start_cell = rowcol_to_a1(target_row, 1)
        end_cell = rowcol_to_a1(target_row, header_len)
        sheet.update(f"{start_cell}:{end_cell}", [
                     new_row], value_input_option="USER_ENTERED")


def buscar_afiliado_por_ruc_base_datos(ruc):
    """Busca un afiliado únicamente en la hoja SOCIOS."""
    ruc = limpiar_ruc(ruc)
    sheet = _get_base_datos_sheet()
    rows = _get_all_records_flexible(sheet, head=2)
    for row in rows:
        if limpiar_ruc(row.get("RUC", "")) == ruc:
            return {
                "razon_social": row.get("RAZON_SOCIAL", ""),
                "ciudad": row.get("CIUDAD", ""),
                "fecha_afiliacion": row.get("FECHA_AFILIACION", ""),
            }
    return None


def obtener_ventas_por_ruc(ruc):
    """Obtiene ventas históricas del afiliado desde VENTAS_SOCIO y, si no hay, desde columnas por año en SOCIOS."""
    ruc_norm = limpiar_ruc(ruc)
    if not ruc_norm:
        return []

    sheet_id = os.getenv("SHEET_PATH") or getattr(settings, "SHEET_PATH", "")
    if not sheet_id:
        return []

    try:
        sheet = get_google_sheet(sheet_id, "VENTAS_SOCIO")
    except Exception:
        return []

    rows = _get_all_records_flexible(sheet, head=2)
    ventas = []
    for row in rows:
        if limpiar_ruc(row.get("RUC", "")) != ruc_norm:
            continue
        anio = str(row.get("ANIO") or row.get("AÑO")
                   or row.get("ANO") or "").strip()
        comparativo = row.get("COMPARATIVO", "")
        ventas_estimadas = (
            row.get("VENTAS_ESTIMADAS")
            or row.get("MONTO_ESTIMADO")
            or row.get("MONTO_VENTAS")
            or row.get("VENTAS_ESTIMADA")
            or ""
        )
        fecha_registro = row.get("FECHA_REGISTRO", "") or row.get("FECHA", "")
        ventas.append(
            {
                "anio": anio,
                "comparativo": comparativo,
                "ventas_estimadas": ventas_estimadas,
                "fecha_registro": fecha_registro,
            }
        )

    # Fallback: buscar columnas por año (ej. 2019, 2020) en la hoja SOCIOS
    try:
        base_sheet = get_google_sheet(sheet_id, "SOCIOS")
        base_rows = _get_all_records_flexible(base_sheet, head=2)
    except Exception:
        base_rows = []

    if base_rows:
        try:
            base_row = next(
                (row for row in base_rows if limpiar_ruc(
                    row.get("RUC", "")) == ruc_norm),
                None,
            )
        except Exception:
            base_row = None
        if base_row:
            existing_years = {v.get("anio") for v in ventas if v.get("anio")}
            for key, value in base_row.items():
                key_str = (key or "").strip()
                if not key_str or not re.fullmatch(r"\d{4}", key_str):
                    continue
                if key_str in existing_years:
                    continue
                val_str = (value or "").strip() if isinstance(
                    value, str) else value
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

    # Ordenar desc por año si es numérico
    ventas.sort(key=lambda v: v.get("anio") or "", reverse=True)
    return ventas


def guardar_ventas_afiliado(data: Dict[str, str]):
    """
    Inserta un registro en la hoja VENTAS_SOCIO con el orden exacto:
    RUC | RAZON_SOCIAL | CIUDAD | FECHA_AFILIACION | REGISTRO_VENTAS |
    COMPARATIVO | MONTO_ESTIMADO | OBSERVACIONES | FECHA_REGISTRO | ANIO
    """
    logging.info("Datos recibidos para guardar ventas: %s", data)

    sheet_id = os.getenv("SHEET_PATH") or getattr(settings, "SHEET_PATH", "")
    if not sheet_id:
        raise RuntimeError("SHEET_PATH no esta configurado.")

    sheet = get_google_sheet(sheet_id, "VENTAS_SOCIO")

    fila = [
        data.get("ruc", ""),
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

    # Inserta asegurando que se respeten las primeras columnas (A-J) en la siguiente fila disponible
    next_row = find_first_empty_row(sheet, start_row=2)
    start = f"A{next_row}"
    end = f"J{next_row}"
    sheet.update(f"{start}:{end}", [fila], value_input_option="USER_ENTERED")
    try:
        # Forzar formato de fecha dd/MM/YYYY en la columna D a partir de la fila 2
        sheet.format(f"D2:D{next_row}", {"numberFormat": {
                     "type": "DATE", "pattern": "dd/MM/yyyy"}})
    except Exception:
        logging.warning(
            "No se pudo aplicar formato de fecha a la columna D en VENTAS_SOCIO.")
