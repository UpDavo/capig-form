from django.shortcuts import render, redirect
from django.contrib import messages
from django.views.decorators.http import require_http_methods
from django.views.decorators.http import require_GET
from django.conf import settings
from django.utils.timezone import now
from datetime import datetime
import pytz
import json
import re

from capig_form.services.google_sheets_service import (
    _get_client,
    get_column_data,
    get_google_sheet,
    insert_row_to_sheet,
)
from forms.afiliacion_handler import guardar_nuevo_afiliado_en_google_sheets
from forms.utils import (
    buscar_afiliado_por_ruc,
    actualizar_estado_afiliado,
    buscar_afiliado_por_ruc_base_datos,
    guardar_ventas_afiliado,
    obtener_ventas_por_ruc,
)

VENTA_KEY_PATTERN = re.compile(r"ventas\[(\d+)\]\[(\w+)\]")


def _entrada_venta_vacia():
    """Estructura base para renderizar un bloque de ventas."""
    return {"anio": "", "comparativo": "", "ventas_estimadas": ""}


def _parsear_bloques_ventas(post_data):
    """Extrae los bloques enviados como ventas[n][campo] preservando el orden."""
    bloques = {}
    for key, value in post_data.items():
        match = VENTA_KEY_PATTERN.fullmatch(key)
        if not match:
            continue
        idx, campo = match.groups()
        campo_normalizado = "comparativo" if campo == "comparar" else campo
        bloques.setdefault(int(idx), {})[campo_normalizado] = (value or "").strip()
    return [bloques[i] for i in sorted(bloques)]

def home_view(request):
    """Vista de inicio con opciones"""
    return render(request, 'home.html')


def _obtener_sectores():
    """Devuelve la lista de sectores desde la hoja 'SECTOR' (columna A, desde A2)."""
    try:
        sheet = get_google_sheet(settings.SHEET_PATH, "SECTOR")
    except Exception:
        # Intentar encontrar la hoja por nombre, aunque tenga espacios o diferencias de may?sculas/min?sculas
        try:
            client = _get_client()
            spreadsheet = client.open_by_key(settings.SHEET_PATH)
            sheet = next((ws for ws in spreadsheet.worksheets() if ws.title.strip().lower() == "sector"), None)
            if not sheet:
                return []
        except Exception:
            return []

    try:
        valores = sheet.col_values(1)
        # Saltar encabezado (fila 1) y limpiar vac?os
        sectores = [val.strip() for val in valores[1:] if val.strip()]
        return sectores
    except Exception:
        return []


def _codigo_seguridad_valido(request):
    """Valida el código de seguridad de 6 dígitos enviado en el POST."""
    codigo = (request.POST.get("security_code") or "").strip()
    return codigo and codigo == getattr(settings, "SECURITY_CODE", "")


def _to_iso_date(fecha_str: str) -> str:
    """
    Intenta convertir fechas como 'DD/MM/YYYY' o 'YYYY-MM-DD' a 'YYYY-MM-DD'.
    Si falla, retorna la cadena original.
    """
    if not fecha_str:
        return fecha_str
    fecha_str = fecha_str.strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(fecha_str, fmt).date().isoformat()
        except ValueError:
            continue
    return fecha_str


@require_http_methods(["GET", "POST"])
def diag_form_view(request):
    """Vista para el formulario de diagnóstico"""
    if request.method == "POST":
        # Nombre de hoja actualizado
        SHEET_NAME = 'ASESORIAS'

        razon_social = request.POST.get('razon_social')
        tipo_diagnostico = request.POST.get('tipo_diagnostico')
        subtipo_diagnostico = request.POST.get('subtipo_diagnostico', '')
        otros_subtipo = request.POST.get('otros_subtipo', '')
        se_diagnostico = request.POST.get('se_diagnostico') == 'true'

        ecuador_tz = pytz.timezone('America/Guayaquil')
        now_ecuador = datetime.now(ecuador_tz)
        fecha_str = now_ecuador.strftime('%Y-%m-%d')
        hora_str = now_ecuador.strftime('%H:%M:%S')

        print("\n=== FORMULARIO DE ASESORÍAS ===")
        print(f"Razón Social: {razon_social}")
        print(f"Tipo de Asesoría: {tipo_diagnostico}")
        print(f"Subtipo de Asesoría: {subtipo_diagnostico}")
        print(f"Otros Subtipo: {otros_subtipo}")
        print(f"Se realizó la asesoría: {se_diagnostico}")
        print(f"Fecha: {fecha_str}")
        print(f"Hora: {hora_str}")
        print("================================\n")

        success = insert_row_to_sheet(settings.SHEET_PATH, SHEET_NAME, [
            razon_social,
            tipo_diagnostico,
            subtipo_diagnostico,
            otros_subtipo,
            'Sí' if se_diagnostico else 'No',
            fecha_str,
            hora_str,
        ])

        if success:
            return redirect('forms:success')
        else:
            messages.error(
                request, 'Hubo un error al guardar los datos. Por favor, intente nuevamente.')

    empresas = get_column_data(
        settings.SHEET_PATH, worksheet_index=0, column='B', start_row=3)

    if not empresas:
        empresas = [
            'Empresa Ejemplo S.A.',
            'Corporación ABC Ltda.',
            'Inversiones XYZ S.A.S.',
            'Grupo Empresarial 123',
            'Soluciones Tecnológicas DEF'
        ]

    return render(request, 'diag_form.html', {'empresas': empresas})


@require_http_methods(["GET", "POST"])
def cap_form_view(request):
    """Vista para el formulario de capacitación"""
    if request.method == "POST":
        SHEET_NAME = 'CAPACITACIONES'

        razon_social = request.POST.get('razon_social')
        nombre_capacitacion = request.POST.get('nombre_capacitacion')
        tipo_capacitacion = request.POST.get('tipo_capacitacion')
        valor_pago = request.POST.get('valor_pago')

        ecuador_tz = pytz.timezone('America/Guayaquil')
        now_ecuador = datetime.now(ecuador_tz)
        fecha_str = now_ecuador.strftime('%Y-%m-%d')
        hora_str = now_ecuador.strftime('%H:%M:%S')

        print("\n=== FORMULARIO DE CAPACITACIÓN ===")
        print(f"Razón Social: {razon_social}")
        print(f"Nombre de la Capacitación: {nombre_capacitacion}")
        print(f"Tipo de Capacitación: {tipo_capacitacion}")
        print(f"Valor del Pago: ${valor_pago}")
        print(f"Fecha: {fecha_str}")
        print(f"Hora: {hora_str}")
        print("==================================\n")

        success = insert_row_to_sheet(settings.SHEET_PATH, SHEET_NAME, [
            razon_social,
            nombre_capacitacion,
            tipo_capacitacion,
            valor_pago,
            fecha_str,
            hora_str,
        ])

        if success:
            return redirect('forms:success')
        else:
            messages.error(
                request, 'Hubo un error al guardar los datos. Por favor, intente nuevamente.')

    empresas = get_column_data(
        settings.SHEET_PATH, worksheet_index=0, column='B', start_row=3)

    if not empresas:
        empresas = [
            'Empresa Ejemplo S.A.',
            'Corporación ABC Ltda.',
            'Inversiones XYZ S.A.S.',
            'Grupo Empresarial 123',
            'Soluciones Tecnológicas DEF'
        ]

    return render(request, 'cap_form.html', {'empresas': empresas})


def success_view(request):
    """Vista de exito despues de enviar el formulario"""
    return render(request, 'success.html')

@require_GET
def success_afiliado_view(request):
    """Vista de exito especifica para afiliacion"""
    return render(request, 'success_afiliado.html')

def custom_404_view(request, exception):
    """Vista personalizada para error 404"""
    return render(request, '404.html', status=404)

@require_GET
def registro_inicio_view(request):
    """Landing para acceso al registro de afiliados."""
    return render(request, "registro_inicio.html")


@require_GET
def estado_inicio_view(request):
    """Landing para el flujo de estado de afiliado."""
    return render(request, "estado_inicio.html")


@require_http_methods(["GET", "POST"])
def estado_afiliado_view(request):
    """Consulta y actualiza el estado de un afiliado."""
    context = {}

    if request.method == "POST":
        ruc = request.POST.get("ruc")
        nuevo_estado = request.POST.get("estado")

        afiliado = buscar_afiliado_por_ruc(ruc)

        if afiliado:
            if nuevo_estado:
                if not _codigo_seguridad_valido(request):
                    messages.error(request, "Código de seguridad inválido.")
                    context["afiliado"] = afiliado
                    return render(request, "estado_afiliado.html", context)
                actualizar_estado_afiliado(ruc, nuevo_estado)
                return redirect("forms:success_estado_afiliado")
            context["afiliado"] = afiliado
        else:
            context["no_encontrado"] = True

    return render(request, "estado_afiliado.html", context)


@require_GET
def success_estado_afiliado_view(request):
    """Confirmación de actualización de estado."""
    return render(request, "success_estado_afiliado.html")

@require_http_methods(["GET", "POST"])
def nuevo_afiliado_view(request):
    """Formulario para registrar un nuevo afiliado en la hoja BASE DE DATOS."""
    sectores = _obtener_sectores()

    if request.method == "POST":
        if not _codigo_seguridad_valido(request):
            messages.error(request, "Código de seguridad inválido.")
            return render(request, "afiliado_form.html", {"sectores": sectores})

        razon_social = request.POST.get("razon_social", "").strip()
        ruc = request.POST.get("ruc", "").strip()
        ciudad = request.POST.get("ciudad", "").strip()
        direccion = request.POST.get("direccion", "").strip()
        telefono = request.POST.get("telefono", "").strip()
        email = request.POST.get("email", "").strip()
        representante = request.POST.get("representante", "").strip()
        cargo = request.POST.get("cargo", "").strip()
        genero = request.POST.get("genero", "").strip()
        colaboradores = request.POST.get("colaboradores", "").strip()
        sector = request.POST.get("sector", "").strip()
        tamano = request.POST.get("tamano", "").strip()
        estado = request.POST.get("estado", "").strip()

        # Fecha actual en zona horaria Guayaquil
        guayaquil = pytz.timezone("America/Guayaquil")
        fecha_afiliacion = now().astimezone(guayaquil).date().isoformat()

        try:
            guardar_nuevo_afiliado_en_google_sheets({
                "razon_social": razon_social,
                "ruc": ruc,
                "fecha_afiliacion": fecha_afiliacion,
                "ciudad": ciudad,
                "direccion": direccion,
                "telefono": telefono,
                "email": email,
                "representante": representante,
                "cargo": cargo,
                "genero": genero,
                "colaboradores": colaboradores,
                "sector": sector,
                "tamano": tamano,
                "estado": estado,
            })
            messages.success(request, "Afiliado registrado correctamente.")
        except Exception as exc:
            messages.error(request, f"Error al registrar: {exc}")

    return render(request, "afiliado_form.html", {"sectores": sectores})


@require_GET
def ventas_inicio_view(request):
    """Landing para el flujo de registro de ventas de afiliados."""
    return render(request, "ventas_inicio.html")


@require_http_methods(["GET", "POST"])
def ventas_afiliado_view(request):
    """Formulario para registrar las ventas de un afiliado (busqueda y envio separados)."""
    context = {
        "ventas_data": [_entrada_venta_vacia()],
        "ruc": request.POST.get("ruc", "").strip(),
        "registro_ventas": request.POST.get("registro_ventas", "").strip(),
        "observaciones": request.POST.get("observaciones", "").strip(),
        "ventas_previas": [],
        "ventas_previas_years": [],
        "ventas_previas_json": "[]",
    }

    if request.method == "POST":
        ruc = request.POST.get("ruc", "").strip()
        registro_ventas = request.POST.get("registro_ventas")
        observaciones = request.POST.get("observaciones", "").strip()
        ventas_bloques = _parsear_bloques_ventas(request.POST)

        afiliado = buscar_afiliado_por_ruc_base_datos(ruc)

        if afiliado:
            ventas_previas = obtener_ventas_por_ruc(ruc)
            years_previas = sorted(
                {v["anio"] for v in ventas_previas if v.get("anio")},
                reverse=True,
            )
            context["afiliado"] = afiliado
            context["ruc"] = ruc
            context["registro_ventas"] = registro_ventas
            context["ventas_data"] = ventas_bloques or [_entrada_venta_vacia()]
            context["observaciones"] = observaciones
            context["ventas_previas"] = ventas_previas
            context["ventas_previas_years"] = years_previas
            context["ventas_previas_json"] = json.dumps(ventas_previas, ensure_ascii=False)

            # Fase 1: solo se busco por RUC, aun no se responde el formulario
            if not registro_ventas:
                return render(request, "ventas_afiliado.html", context)

            # Fase 2: ya respondio preguntas -> guardar
            if not _codigo_seguridad_valido(request):
                messages.error(request, "Codigo de seguridad invalido.")
                return render(request, "ventas_afiliado.html", context)

            rv_norm = (registro_ventas or "").strip().lower()
            es_si = rv_norm in {"si", "sí", "s\u00ed", "s"}

            base_data = {
                "ruc": ruc,
                "razon_social": afiliado["razon_social"],
                "ciudad": afiliado["ciudad"],
                # Enviamos ISO para que Sheets lo interprete como fecha, el formato de visualizacion se manejara en la hoja.
                "fecha_afiliacion": _to_iso_date(afiliado["fecha_afiliacion"]),
                "registro_ventas": registro_ventas,
                "observaciones": observaciones,
                "fecha_registro": datetime.now().strftime("%Y-%m-%d %H:%M"),
            }

            if es_si:
                if not ventas_bloques:
                    messages.error(request, "Agrega al menos un registro de ventas anual.")
                    return render(request, "ventas_afiliado.html", context)

                for bloque in ventas_bloques:
                    anio = (bloque.get("anio") or "").strip()
                    if not anio:
                        messages.error(request, "Selecciona el anio para cada registro de ventas.")
                        return render(request, "ventas_afiliado.html", context)

                    data = {
                        **base_data,
                        "comparativo": bloque.get("comparativo", ""),
                        "ventas_estimadas": bloque.get("ventas_estimadas", ""),
                        "anio": anio,
                    }
                    guardar_ventas_afiliado(data)
            else:
                data = {
                    **base_data,
                    "comparativo": "",
                    "ventas_estimadas": "",
                    "anio": str(datetime.now().year),
                }
                guardar_ventas_afiliado(data)
            return redirect("forms:success_ventas_afiliado")
        else:
            context["no_encontrado"] = True

    return render(request, "ventas_afiliado.html", context)

@require_GET
def success_ventas_afiliado_view(request):
    """Confirmacion de registro de ventas."""
    return render(request, "success_ventas_afiliado.html")
