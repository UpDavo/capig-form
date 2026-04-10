import json
import logging
import re
from datetime import datetime, timedelta

import pytz
from django.conf import settings
from django.contrib import messages
from django.shortcuts import redirect, render
from django.utils.timezone import now
from django.views.decorators.http import require_GET, require_http_methods

from capig_form.services.google_sheets_service import (
    _get_client,
    get_google_sheet,
    insert_row_to_sheet,
)
from forms.afiliacion_handler import (
    EMAIL_COLUMN_SEQUENCE,
    PHONE_COLUMN_SEQUENCE,
    guardar_nuevo_afiliado_en_google_sheets,
)
from forms.utils import (
    actualizar_estado_afiliado,
    buscar_afiliado_por_ruc,
    buscar_afiliado_por_ruc_base_datos,
    guardar_ventas_afiliado,
    listar_empresas_socias,
    limpiar_ruc,
    obtener_ventas_por_ruc,
)

logger = logging.getLogger(__name__)

VENTA_KEY_PATTERN = re.compile(r"ventas\[(\d+)\]\[(\w+)\]")
AFILIADO_FORM_FIELDS = [
    "razon_social",
    "ruc",
    "ciudad",
    "direccion",
    "actividad",
    "representante",
    "cargo",
    "gerente_general_nombre",
    "genero",
    "gerente_general_contacto",
    "gerente_general_email",
    "gerente_fin_nombre",
    "gerente_fin_contacto",
    "gerente_fin_email",
    "gerente_rrhh_nombre",
    "gerente_rrhh_contacto",
    "gerente_rrhh_email",
    "gerente_comercial_nombre",
    "gerente_comercial_contacto",
    "gerente_comercial_email",
    "gerente_produccion_nombre",
    "gerente_produccion_contacto",
    "gerente_produccion_email",
    "colaboradores",
    "sector",
    "tamano",
    "estado",
]
AFILIADO_LIST_FIELDS = ["telefonos", "emails"]
AFILIADO_EXECUTIVE_FIELDS = [
    "gerente_general_nombre",
    "genero",
    "gerente_general_contacto",
    "gerente_general_email",
    "gerente_fin_nombre",
    "gerente_fin_contacto",
    "gerente_fin_email",
    "gerente_rrhh_nombre",
    "gerente_rrhh_contacto",
    "gerente_rrhh_email",
    "gerente_comercial_nombre",
    "gerente_comercial_contacto",
    "gerente_comercial_email",
    "gerente_produccion_nombre",
    "gerente_produccion_contacto",
    "gerente_produccion_email",
]


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


def _build_afiliado_form_data(post_data=None):
    """Construye el estado del formulario de afiliado con defaults y datos enviados."""
    data = {field: "" for field in AFILIADO_FORM_FIELDS}
    data["telefonos"] = [""]
    data["emails"] = [""]
    data["estado"] = "PAGADO"
    if not post_data:
        return data

    for field in AFILIADO_FORM_FIELDS:
        data[field] = (post_data.get(field, "") or "").strip()

    for field_name in AFILIADO_LIST_FIELDS:
        values = []
        if hasattr(post_data, "getlist"):
            values = [(value or "").strip() for value in post_data.getlist(field_name)]
        values = values or [""]
        data[field_name] = values

    if not data["estado"]:
        data["estado"] = "PAGADO"
    return data


def _build_afiliado_form_context(sectores, form_data=None):
    data = form_data or _build_afiliado_form_data()
    return {
        "sectores": sectores,
        "form_data": data,
        "max_phone_fields": len(PHONE_COLUMN_SEQUENCE),
        "max_email_fields": len(EMAIL_COLUMN_SEQUENCE),
        "show_executive_contacts": any(data.get(field) for field in AFILIADO_EXECUTIVE_FIELDS),
    }


def dashboard_view(request):
    """Vista del dashboard principal."""
    return render(request, "dashboard.html")


def _obtener_sectores():
    """Devuelve la lista de sectores desde la hoja SECTOR."""
    try:
        sheet = get_google_sheet(settings.SHEET_PATH, "SECTOR")
    except Exception:
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
        return [val.strip() for val in valores[1:] if val.strip()]
    except Exception:
        return []


def _codigo_seguridad_valido(request):
    """Valida el codigo de seguridad enviado en el POST."""
    codigo = (request.POST.get("security_code") or "").strip()
    return codigo and codigo == getattr(settings, "SECURITY_CODE", "")


def _cargar_empresas_socias():
    """Obtiene empresas desde SOCIOS y un set para validacion rapida."""
    empresas = listar_empresas_socias()
    empresas_validas = {
        (empresa.get("razon_social") or "").strip().casefold()
        for empresa in empresas
        if (empresa.get("razon_social") or "").strip()
    }
    return empresas, empresas_validas


def _to_iso_date(fecha_str: str) -> str:
    """
    Intenta convertir fechas como DD/MM/YYYY o YYYY-MM-DD a YYYY-MM-DD.
    Si recibe un serial de Excel, tambien lo convierte.
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
    """Vista para el formulario de diagnostico."""
    empresas, empresas_validas = _cargar_empresas_socias()
    selected_razon_social = (request.POST.get("razon_social", "") or "").strip()

    if request.method == "POST":
        sheet_name = "ASESORIAS"

        razon_social = selected_razon_social
        tipo_diagnostico = request.POST.get("tipo_diagnostico")
        subtipo_diagnostico = request.POST.get("subtipo_diagnostico", "")
        otros_subtipo = request.POST.get("otros_subtipo", "")
        se_diagnostico = request.POST.get("se_diagnostico") == "true"

        if not empresas:
            messages.error(
                request,
                "No se pudo cargar la lista de empresas socias desde Google Sheets. Revisa la hoja SOCIOS.",
            )
        elif razon_social.casefold() not in empresas_validas:
            messages.error(
                request,
                "La empresa seleccionada no pertenece a la lista actual de socios.",
            )
        else:
            ecuador_tz = pytz.timezone("America/Guayaquil")
            now_ecuador = datetime.now(ecuador_tz)
            fecha_str = now_ecuador.strftime("%Y-%m-%d")
            hora_str = now_ecuador.strftime("%H:%M:%S")

            success = insert_row_to_sheet(
                settings.SHEET_PATH,
                sheet_name,
                [
                    razon_social,
                    tipo_diagnostico,
                    subtipo_diagnostico,
                    otros_subtipo,
                    "Si" if se_diagnostico else "No",
                    fecha_str,
                    hora_str,
                ],
            )

            if success:
                return redirect("forms:success")
            messages.error(request, "Hubo un error al guardar los datos. Por favor, intente nuevamente.")

    return render(
        request,
        "diag_form.html",
        {
            "empresas": empresas,
            "selected_razon_social": selected_razon_social,
            "selected_tipo_diagnostico": (request.POST.get("tipo_diagnostico", "") or "").strip(),
            "selected_subtipo_diagnostico": (request.POST.get("subtipo_diagnostico", "") or "").strip(),
            "otros_subtipo_value": (request.POST.get("otros_subtipo", "") or "").strip(),
            "se_diagnostico_checked": request.POST.get("se_diagnostico") == "true",
            "submit_disabled": not empresas,
        },
    )


@require_http_methods(["GET", "POST"])
def cap_form_view(request):
    """Vista para el formulario de capacitacion."""
    empresas, empresas_validas = _cargar_empresas_socias()
    razon_social = (request.POST.get("razon_social", "") or "").strip()
    no_en_lista_checked = request.POST.get("no_en_lista") == "on"

    if request.method == "POST":
        sheet_name = "CAPACITACIONES"

        nombre_capacitacion = request.POST.get("nombre_capacitacion")
        tipo_capacitacion = request.POST.get("tipo_capacitacion")
        valor_pago = request.POST.get("valor_pago")

        if not no_en_lista_checked:
            if not empresas:
                messages.error(
                    request,
                    "No se pudo cargar la lista de empresas socias desde Google Sheets. Marca 'No esta en la lista' solo si corresponde registrar un externo.",
                )
            elif razon_social.casefold() not in empresas_validas:
                messages.error(
                    request,
                    "La empresa seleccionada no pertenece a la lista actual de socios.",
                )
            else:
                ecuador_tz = pytz.timezone("America/Guayaquil")
                now_ecuador = datetime.now(ecuador_tz)
                fecha_str = now_ecuador.strftime("%Y-%m-%d")
                hora_str = now_ecuador.strftime("%H:%M:%S")

                success = insert_row_to_sheet(
                    settings.SHEET_PATH,
                    sheet_name,
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
        else:
            ecuador_tz = pytz.timezone("America/Guayaquil")
            now_ecuador = datetime.now(ecuador_tz)
            fecha_str = now_ecuador.strftime("%Y-%m-%d")
            hora_str = now_ecuador.strftime("%H:%M:%S")

            success = insert_row_to_sheet(
                settings.SHEET_PATH,
                sheet_name,
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

    return render(
        request,
        "cap_form.html",
        {
            "empresas": empresas,
            "selected_razon_social": razon_social,
            "no_en_lista_checked": no_en_lista_checked,
            "nombre_capacitacion_value": (request.POST.get("nombre_capacitacion", "") or "").strip(),
            "tipo_capacitacion_value": (request.POST.get("tipo_capacitacion", "") or "").strip(),
            "valor_pago_value": (request.POST.get("valor_pago", "") or "").strip(),
        },
    )


def success_view(request):
    """Vista de exito despues de enviar el formulario."""
    return render(request, "success.html")


@require_GET
def success_afiliado_view(request):
    """Vista de exito especifica para afiliacion."""
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
    """Confirmacion de actualizacion de estado."""
    estado_update = request.session.pop("estado_update", None)
    context = {"estado_update": estado_update}
    return render(request, "success_estado_afiliado.html", context)


@require_http_methods(["GET", "POST"])
def nuevo_afiliado_view(request):
    """Formulario para registrar un nuevo afiliado en la hoja SOCIOS."""
    sectores = _obtener_sectores()
    form_data = _build_afiliado_form_data(
        request.POST if request.method == "POST" else None
    )
    context = _build_afiliado_form_context(sectores, form_data)

    if request.method == "POST":
        ruc_norm = limpiar_ruc(form_data.get("ruc", ""))
        guayaquil = pytz.timezone("America/Guayaquil")
        fecha_afiliacion = now().astimezone(guayaquil).date().isoformat()

        if len(ruc_norm) != 13:
            context["ruc_error"] = "RUC invalido. Debe tener 13 digitos."
        else:
            try:
                guardar_nuevo_afiliado_en_google_sheets(
                    {
                        **form_data,
                        "ruc": ruc_norm,
                        "fecha_afiliacion": fecha_afiliacion,
                    }
                )
                return redirect("forms:success_afiliado")
            except Exception:  # pragma: no cover - depende de API externa
                logger.exception("Error al registrar afiliado con RUC %s", ruc_norm)
                context["form_error"] = "No se pudo registrar el afiliado. Intenta nuevamente."

    return render(request, "afiliado_form.html", context)


@require_http_methods(["GET", "POST"])
def ventas_afiliado_view(request):
    """Formulario para registrar las ventas de un afiliado."""
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
            es_si = rv_norm in {"si", "si", "s\u00ed", "s"}

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
                        messages.error(request, "Selecciona el ano para cada registro de ventas.")
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
