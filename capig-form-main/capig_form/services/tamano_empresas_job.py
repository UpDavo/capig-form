import os
import unicodedata
from collections import defaultdict
from typing import Dict, List, Tuple

import pandas as pd


def _ensure_django():
    import django
    from django.conf import settings

    if not settings.configured:
        os.environ.setdefault("DJANGO_SETTINGS_MODULE", "capig_form.settings")
        django.setup()


def _normalize_label(value: str) -> str:
    text = str(value or "").strip().upper()
    text = unicodedata.normalize("NFD", text)
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    text = text.replace(" ", "_")
    return text


def _find_col(headers: List[str], needle: str) -> int:
    target = _normalize_label(needle)
    for idx, h in enumerate(headers):
        if target in _normalize_label(h):
            return idx
    return -1


def _clean_ruc(raw) -> str:
    return str(raw or "").replace("'", "").replace("\"", "").strip()


def _parse_year(value) -> int:
    ts = pd.to_datetime(value, errors="coerce", dayfirst=True)
    if pd.isna(ts):
        return None
    return int(ts.year)


def _clasificar_tamano(monto: float) -> str:
    if monto <= 100_000:
        return "MICRO"
    if monto <= 1_000_000:
        return "PEQUENA"
    if monto <= 5_000_000:
        return "MEDIANA"
    return "GRANDE"


def _collect_historicos(data_bd: List[List[str]]) -> List[Tuple[str, int, str]]:
    historicos: List[Tuple[str, int, str]] = []

    def _process_block(header_row: List[str], rows: List[List[str]]):
        col_ruc = _find_col(header_row, "RUC")
        col_tam = _find_col(header_row, "TAMANO")
        col_fecha = _find_col(header_row, "FECHA_AFILIACION")
        if min(col_ruc, col_tam, col_fecha) < 0:
            return
        for row in rows:
            ruc = _clean_ruc(row[col_ruc]) if col_ruc < len(row) else ""
            tam = str(row[col_tam] or "").strip().upper() if col_tam < len(row) else ""
            anio = _parse_year(row[col_fecha]) if col_fecha < len(row) else None
            if ruc and tam and anio:
                historicos.append((ruc, anio, tam))

    # Bloque 1: filas 3-312, encabezado en fila 2 (indices 2:312 y header 1)
    if len(data_bd) >= 2:
        headers1 = data_bd[1]
        bloque1 = data_bd[2:312]
        _process_block(headers1, bloque1)

    # Bloque 2: desde fila 315, encabezado en fila 314 (indices 314: y header 313)
    if len(data_bd) >= 314:
        headers2 = data_bd[313]
        bloque2 = data_bd[314:]
        _process_block(headers2, bloque2)

    return historicos


def _collect_ventas(data_ventas: List[List[str]]) -> Dict[Tuple[str, int], float]:
    if not data_ventas:
        return {}
    headers = data_ventas[0]
    col_ruc = _find_col(headers, "RUC")
    col_monto = _find_col(headers, "MONTO")
    col_anio = _find_col(headers, "ANO")
    if col_anio == -1:
        col_anio = _find_col(headers, "ANIO")
    if min(col_ruc, col_monto, col_anio) < 0:
        return {}
    ventas_agrupadas: Dict[Tuple[str, int], float] = defaultdict(float)
    for row in data_ventas[1:]:
        ruc = _clean_ruc(row[col_ruc]) if col_ruc < len(row) else ""
        anio_raw = row[col_anio] if col_anio < len(row) else None
        try:
            anio = int(anio_raw)
        except (TypeError, ValueError):
            anio = None
        try:
            monto = float(row[col_monto]) if col_monto < len(row) and str(row[col_monto]).strip() else 0.0
        except (TypeError, ValueError):
            monto = 0.0
        if not ruc or not anio:
            continue
        ventas_agrupadas[(ruc, anio)] += monto
    return ventas_agrupadas


def _build_registros(historicos: List[Tuple[str, int, str]], ventas_agrupadas: Dict[Tuple[str, int], float]):
    registros: Dict[str, Dict[int, str]] = defaultdict(dict)
    for ruc, anio, tam in historicos:
        registros[ruc][anio] = tam
    for (ruc, anio), monto in ventas_agrupadas.items():
        registros[ruc][anio] = _clasificar_tamano(monto)
    return registros


def _build_cambios_y_resumen(registros: Dict[str, Dict[int, str]]):
    cambios = []
    resumen: Dict[str, int] = defaultdict(int)
    for ruc, data in registros.items():
        if not data:
            continue
        historial = sorted(
            ((anio, tam) for anio, tam in data.items() if anio and tam),
            key=lambda x: x[0],
        )
        if len(historial) < 2:
            continue
        # Contar transiciones consecutivas año a año
        for (anio_a, tam_a), (anio_b, tam_b) in zip(historial, historial[1:]):
            if tam_a == tam_b:
                continue
            cambios.append([ruc, anio_a, tam_a, anio_b, tam_b])
            clave = f"{tam_a} -> {tam_b}"
            resumen[clave] += 1
    return cambios, resumen


def _build_tamano_global(registros: Dict[str, Dict[int, str]]):
    salida = []
    for ruc, data in registros.items():
        if not data:
            continue
        ultimo_anio = max(data.keys())
        tam = data[ultimo_anio]
        salida.append([ruc, tam])
    return salida


def _ensure_worksheet(spreadsheet, name: str):
    from gspread.exceptions import WorksheetNotFound

    try:
        return spreadsheet.worksheet(name)
    except WorksheetNotFound:
        return spreadsheet.add_worksheet(title=name, rows=2000, cols=20)


def _update_sheet(ws, rows: List[List]):
    ws.clear()
    if rows:
        ws.update(rows)


def _tamano_to_code(tamano: str) -> str:
    """
    Convierte nombre de tamaño a código numérico 1-4.
    Retorna string vacío si no coincide.
    """
    mapping = {
        "MICRO": "1",
        "PEQUENA": "2",
        "MEDIANA": "3",
        "GRANDE": "4",
    }
    return mapping.get(tamano.upper().strip(), "")


def _write_t202x_columns(data_bd: List[List[str]], registros: Dict[str, Dict[int, str]]) -> List[List[str]]:
    """
    Actualiza data_bd con columnas T202x basado en los registros calculados.
    Retorna una copia modificada de data_bd sin alterar el original.
    
    Parámetros:
    - data_bd: datos completos de BASE DE DATOS (tal como se leen de Sheets)
    - registros: dict {RUC: {anio: tamano}} ya calculado
    
    Retorna:
    - data_bd actualizada con columnas T202x
    """
    if not data_bd or len(data_bd) < 2:
        return data_bd
    
    # Trabajar con copia para no modificar el original
    data_bd_copy = [row[:] for row in data_bd]
    
    # Obtener años únicos de todos los registros
    anios_unicos = set()
    for ruc_data in registros.values():
        anios_unicos.update(ruc_data.keys())
    anios_ordenados = sorted([a for a in anios_unicos if a and isinstance(a, int)])
    
    if not anios_ordenados:
        return data_bd_copy
    
    # Procesar bloque 1 (filas 3-312, header en fila 2)
    if len(data_bd_copy) >= 2:
        headers1 = data_bd_copy[1]
        col_ruc_1 = _find_col(headers1, "RUC")
        
        if col_ruc_1 >= 0:
            # Crear o encontrar columnas T202x
            t_col_indices_1 = {}
            for anio in anios_ordenados:
                col_name = f"T{anio}"
                col_idx = _find_col(headers1, col_name)
                if col_idx == -1:
                    # Agregar nueva columna al header
                    headers1.append(col_name)
                    col_idx = len(headers1) - 1
                t_col_indices_1[anio] = col_idx
            
            # Actualizar filas del bloque 1
            for i in range(2, min(312, len(data_bd_copy))):
                row = data_bd_copy[i]
                # Extender fila si es necesario
                while len(row) < len(headers1):
                    row.append("")
                
                ruc = _clean_ruc(row[col_ruc_1]) if col_ruc_1 < len(row) else ""
                if not ruc or ruc not in registros:
                    continue
                
                # Escribir códigos T202x
                for anio, col_idx in t_col_indices_1.items():
                    if anio in registros[ruc]:
                        tamano = registros[ruc][anio]
                        codigo = _tamano_to_code(tamano)
                        row[col_idx] = codigo
    
    # Procesar bloque 2 (desde fila 315, header en fila 314)
    if len(data_bd_copy) >= 314:
        headers2 = data_bd_copy[313]
        col_ruc_2 = _find_col(headers2, "RUC")
        
        if col_ruc_2 >= 0:
            # Crear o encontrar columnas T202x
            t_col_indices_2 = {}
            for anio in anios_ordenados:
                col_name = f"T{anio}"
                col_idx = _find_col(headers2, col_name)
                if col_idx == -1:
                    # Agregar nueva columna al header
                    headers2.append(col_name)
                    col_idx = len(headers2) - 1
                t_col_indices_2[anio] = col_idx
            
            # Actualizar filas del bloque 2
            for i in range(314, len(data_bd_copy)):
                row = data_bd_copy[i]
                # Extender fila si es necesario
                while len(row) < len(headers2):
                    row.append("")
                
                ruc = _clean_ruc(row[col_ruc_2]) if col_ruc_2 < len(row) else ""
                if not ruc or ruc not in registros:
                    continue
                
                # Escribir códigos T202x
                for anio, col_idx in t_col_indices_2.items():
                    if anio in registros[ruc]:
                        tamano = registros[ruc][anio]
                        codigo = _tamano_to_code(tamano)
                        row[col_idx] = codigo
    
    return data_bd_copy


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

    historicos = _collect_historicos(data_bd)
    ventas_agrupadas = _collect_ventas(data_ventas)
    registros = _build_registros(historicos, ventas_agrupadas)
    cambios, resumen = _build_cambios_y_resumen(registros)
    tamano_global = _build_tamano_global(registros)

    hoja_detalle = _ensure_worksheet(ss, "CAMBIO_TAMANIO_EMPRESAS")
    hoja_resumen = _ensure_worksheet(ss, "RESUMEN_CAMBIOS_TAMANIO")
    hoja_global = _ensure_worksheet(ss, "TAMANO_EMPRESA_GLOBAL")

    detalle_rows = [["RUC", "Ano Inicial", "Tamano Inicial", "Ano Final", "Tamano Final"]] + cambios
    resumen_rows = [["Cambio", "Empresas", "%"]]
    total = len(cambios)
    if total > 0:
        for clave, cuenta in resumen.items():
            resumen_rows.append([clave, cuenta, f"{(cuenta / total) * 100:.2f}%"])
    global_rows = [["RUC", "Tamano"]] + tamano_global

    _update_sheet(hoja_detalle, detalle_rows)
    _update_sheet(hoja_resumen, resumen_rows)
    _update_sheet(hoja_global, global_rows)

    # Actualizar columnas T202x en BASE DE DATOS
    print("[tamano_empresas_job] Actualizando columnas T202x en BASE DE DATOS...")
    data_bd_actualizada = _write_t202x_columns(data_bd, registros)
    if data_bd_actualizada:
        _update_sheet(hoja_bd, data_bd_actualizada)
        print("[tamano_empresas_job] Columnas T202x actualizadas exitosamente.")
    else:
        print("[tamano_empresas_job] No se generaron columnas T202x (datos insuficientes).")

    # Backup local opcional
    output_path = os.path.join(os.path.dirname(__file__), "data", "cambio_tamano_empresas.xlsx")
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    with pd.ExcelWriter(output_path) as writer:
        pd.DataFrame(detalle_rows[1:], columns=detalle_rows[0]).to_excel(writer, sheet_name="cambios", index=False)
        pd.DataFrame(resumen_rows[1:], columns=resumen_rows[0]).to_excel(writer, sheet_name="resumen", index=False)
        pd.DataFrame(global_rows[1:], columns=global_rows[0]).to_excel(writer, sheet_name="tamano_global", index=False)


if __name__ == "__main__":
    run()
