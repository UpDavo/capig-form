import json
import logging
import re
from datetime import datetime, timedelta

import pytz
from django.conf import settings
from django.contrib import messages
from django.shortcuts import render, redirect
from django.utils.timezone import now
from django.views.decorators.http import require_GET, require_http_methods

from capig_form.services.google_sheets_service import (
    _get_client,
    get_column_data,
    get_google_sheet,
    insert_row_to_sheet,
)
from forms.afiliacion_handler import guardar_nuevo_afiliado_en_google_sheets
from forms.utils import (
    actualizar_estado_afiliado,
    buscar_afiliado_por_ruc,
    buscar_afiliado_por_ruc_base_datos,
    guardar_ventas_afiliado,
    limpiar_ruc,
    obtener_ventas_por_ruc,
)

logger = logging.getLogger(__name__)

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


def dashboard_view(request):
    """Vista del dashboard principal (con layout)."""
    return render(request, "dashboard.html")


def _obtener_sectores():
    """Devuelve la lista de sectores desde la hoja 'SECTOR' (columna A, desde A2)."""
    try:
        sheet = get_google_sheet(settings.SHEET_PATH, "SECTOR")
    except Exception:
        # Intentar encontrar la hoja por nombre, aunque tenga espacios o diferencias de mayúsculas/minúsculas
        try:
            client = _get_client()
            spreadsheet = client.open_by_key(settings.SHEET_PATH)
            sheet = next(
                (ws for ws in spreadsheet.worksheets() if ws.title.strip().lower() == "sector"),
                None,
            )
            if not sheet:
                return []
        except Exception:
            return []

    try:
        valores = sheet.col_values(1)
        # Saltar encabezado (fila 1) y limpiar vacíos
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
    Si recibe un serial de Excel (int/float), lo convierte. Si falla, retorna la cadena original.
    """
    if fecha_str is None:
        return fecha_str
    if isinstance(fecha_str, (int, float)):
        try:
            base = datetime(1899, 12, 30)
            return (base + timedelta(days=float(fecha_str))).date().isoformat()
        except Exception:
            return fecha_str
    fecha_str = str(fecha_str).strip()
    if not fecha_str:
        return fecha_str
    for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(fecha_str, fmt).date().isoformat()
        except ValueError:
            continue
    return fecha_str


@require_http_methods(["GET", "POST"])
def diag_form_view(request):
    """Vista para el formulario de diagnóstico."""
    if request.method == "POST":
        SHEET_NAME = "ASESORIAS"

        razon_social = request.POST.get("razon_social")
        tipo_diagnostico = request.POST.get("tipo_diagnostico")
        subtipo_diagnostico = request.POST.get("subtipo_diagnostico", "")
        otros_subtipo = request.POST.get("otros_subtipo", "")
        se_diagnostico = request.POST.get("se_diagnostico") == "true"

        ecuador_tz = pytz.timezone("America/Guayaquil")
        now_ecuador = datetime.now(ecuador_tz)
        fecha_str = now_ecuador.strftime("%Y-%m-%d")
        hora_str = now_ecuador.strftime("%H:%M:%S")

        success = insert_row_to_sheet(
            settings.SHEET_PATH,
            SHEET_NAME,
            [
                razon_social,
                tipo_diagnostico,
                subtipo_diagnostico,
                otros_subtipo,
                "Sí" if se_diagnostico else "No",
                fecha_str,
                hora_str,
            ],
        )

        if success:
            return redirect("forms:success")
        messages.error(request, "Hubo un error al guardar los datos. Por favor, intente nuevamente.")

    empresas = get_column_data(settings.SHEET_PATH, worksheet_index=0, column="B", start_row=3)

    if not empresas:
        empresas = [
            "Empresa Ejemplo S.A.",
            "Corporación ABC Ltda.",
            "Inversiones XYZ S.A.S.",
            "Grupo Empresarial 123",
            "Soluciones Tecnológicas DEF",
        ]

    return render(request, "diag_form.html", {"empresas": empresas})


@require_http_methods(["GET", "POST"])
def cap_form_view(request):
    """Vista para el formulario de capacitación."""
    if request.method == "POST":
        SHEET_NAME = "CAPACITACIONES"

        razon_social = request.POST.get("razon_social")
        nombre_capacitacion = request.POST.get("nombre_capacitacion")
        tipo_capacitacion = request.POST.get("tipo_capacitacion")
        valor_pago = request.POST.get("valor_pago")

        ecuador_tz = pytz.timezone("America/Guayaquil")
        now_ecuador = datetime.now(ecuador_tz)
        fecha_str = now_ecuador.strftime("%Y-%m-%d")
        hora_str = now_ecuador.strftime("%H:%M:%S")

        success = insert_row_to_sheet(
            settings.SHEET_PATH,
            SHEET_NAME,
            [
                razon_social,
                nombre_capacitacion,
                tipo_capacitacion,
                valor_pago,
                fecha_str,
                hora_str,
            ],
        )

        if success:
            return redirect("forms:success")
        messages.error(request, "Hubo un error al guardar los datos. Por favor, intente nuevamente.")

    empresas = get_column_data(settings.SHEET_PATH, worksheet_index=0, column="B", start_row=3)

    if not empresas:
        empresas = [
            "Empresa Ejemplo S.A.",
            "Corporación ABC Ltda.",
            "Inversiones XYZ S.A.S.",
            "Grupo Empresarial 123",
            "Soluciones Tecnológicas DEF",
        ]

    return render(request, "cap_form.html", {"empresas": empresas})


def success_view(request):
    """Vista de éxito después de enviar el formulario."""
    return render(request, "success.html")


@require_GET
def success_afiliado_view(request):
    """Vista de éxito específica para afiliación."""
    return render(request, "success_afiliado.html")


def custom_404_view(request, exception):
    """Vista personalizada para error 404."""
    return render(request, "404.html", status=404)


@require_http_methods(["GET", "POST"])
def estado_afiliado_view(request):
    """Consulta y actualiza el estado de un afiliado."""
    context = {}

    if request.method == "POST":
        ruc = request.POST.get("ruc")
        ruc_norm = limpiar_ruc(ruc)
        nuevo_estado = request.POST.get("estado")

        afiliado = buscar_afiliado_por_ruc(ruc_norm)

        if afiliado:
            if nuevo_estado:
                actualizar_estado_afiliado(ruc_norm, nuevo_estado)
                request.session["estado_update"] = {
                    "razon_social": afiliado.get("razon_social", "N/A"),
                    "ruc": ruc_norm,
                    "estado_anterior": afiliado.get("estado", "N/A"),
                    "estado_nuevo": nuevo_estado,
                }
                return redirect("forms:success_estado_afiliado")
            context["afiliado"] = afiliado
        else:
            context["no_encontrado"] = True
            context["ruc"] = ruc_norm

    return render(request, "estado_afiliado.html", context)


@require_GET
def success_estado_afiliado_view(request):
    """Confirmación de actualización de estado."""
    estado_update = request.session.pop("estado_update", None)
    context = {"estado_update": estado_update}
    return render(request, "success_estado_afiliado.html", context)


@require_http_methods(["GET", "POST"])
def nuevo_afiliado_view(request):
    """Formulario para registrar un nuevo afiliado en la hoja SOCIOS."""
    sectores = _obtener_sectores()
    context = {"sectores": sectores}

    if request.method == "POST":
        razon_social = request.POST.get("razon_social", "").strip()
        ruc = request.POST.get("ruc", "").strip()
        ruc_norm = limpiar_ruc(ruc)
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

        guayaquil = pytz.timezone("America/Guayaquil")
        fecha_afiliacion = now().astimezone(guayaquil).date().isoformat()

        context.update(
            {
                "razon_social": razon_social,
                "ruc": ruc,
                "ciudad": ciudad,
                "direccion": direccion,
                "telefono": telefono,
                "email": email,
                "representante": representante,
                "cargo": cargo,
                "genero": genero,
                "colaboradores": colaboradores,
                "sector_sel": sector,
                "tamano_sel": tamano,
                "estado_sel": estado,
            }
        )

        if len(ruc_norm) != 13:
            context["ruc_error"] = "RUC inválido. Debe tener 13 dígitos."
        else:
            try:
                guardar_nuevo_afiliado_en_google_sheets(
                    {
                        "razon_social": razon_social,
                        "ruc": ruc_norm,
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
                    }
                )
                return redirect("forms:success_afiliado")
            except Exception:  # pragma: no cover - depende de API externa
                logger.exception("Error al registrar afiliado con RUC %s", ruc_norm)
                context["form_error"] = "No se pudo registrar el afiliado. Intenta nuevamente."

    return render(request, "afiliado_form.html", context)


@require_http_methods(["GET", "POST"])
def ventas_afiliado_view(request):
    """Formulario para registrar las ventas de un afiliado (búsqueda y envío separados)."""
    ruc_inicial = limpiar_ruc(request.POST.get("ruc", "").strip())
    context = {
        "ventas_data": [_entrada_venta_vacia()],
        "ruc": ruc_inicial,
        "registro_ventas": request.POST.get("registro_ventas", "").strip(),
        "observaciones": request.POST.get("observaciones", "").strip(),
        "ventas_previas": [],
        "ventas_previas_years": [],
        "ventas_previas_json": "[]",
    }

    if request.method == "POST":
        ruc = request.POST.get("ruc", "").strip()
        ruc_norm = limpiar_ruc(ruc)
        registro_ventas = request.POST.get("registro_ventas")
        observaciones = request.POST.get("observaciones", "").strip()
        ventas_bloques = _parsear_bloques_ventas(request.POST)

        afiliado = buscar_afiliado_por_ruc_base_datos(ruc_norm)

        if afiliado:
            ventas_previas = obtener_ventas_por_ruc(ruc_norm)
            years_previas = sorted({v["anio"] for v in ventas_previas if v.get("anio")}, reverse=True)
            context["afiliado"] = afiliado
            context["ruc"] = ruc_norm
            context["registro_ventas"] = registro_ventas
            context["ventas_data"] = ventas_bloques or [_entrada_venta_vacia()]
            context["observaciones"] = observaciones
            context["ventas_previas"] = ventas_previas
            context["ventas_previas_years"] = years_previas
            context["ventas_previas_json"] = json.dumps(ventas_previas, ensure_ascii=False)

            if not registro_ventas:
                return render(request, "ventas_afiliado.html", context)

            rv_norm = (registro_ventas or "").strip().lower()
            es_si = rv_norm in {"si", "sí", "s\u00ed", "s"}

            base_data = {
                "ruc": ruc_norm,
                "razon_social": afiliado["razon_social"],
                "ciudad": afiliado["ciudad"],
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
                        messages.error(request, "Selecciona el año para cada registro de ventas.")
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
    """Confirmación de registro de ventas."""
    return render(request, "success_ventas_afiliado.html")
