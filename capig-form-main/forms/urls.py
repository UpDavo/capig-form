from django.urls import path
from .view.form_views import home_view, registro_inicio_view, diag_form_view, cap_form_view, success_view, success_afiliado_view, success_estado_afiliado_view, estado_afiliado_view, nuevo_afiliado_view, estado_inicio_view, ventas_inicio_view, ventas_afiliado_view, success_ventas_afiliado_view

app_name = 'forms'

urlpatterns = [
    path("", home_view, name="home"),
    path("registro-inicio/", registro_inicio_view, name="registro_inicio"),
    path('diagnostico/', diag_form_view, name='diag_form'),
    path('capacitacion/', cap_form_view, name='cap_form'),
    path("estado-afiliado/", estado_afiliado_view, name="estado_afiliado"),
    path("exito-estado-afiliado/", success_estado_afiliado_view, name="success_estado_afiliado"),
    path("registrar-afiliado/", nuevo_afiliado_view, name="nuevo_afiliado"),
    path("exito-afiliado/", success_afiliado_view, name="success_afiliado"),
    path('exito/', success_view, name='success'),
    path("estado-inicio/", estado_inicio_view, name="estado_inicio"),
    path("ventas-inicio/", ventas_inicio_view, name="ventas_inicio"),
    path("ventas-afiliado/", ventas_afiliado_view, name="ventas_afiliado"),
    path("exito-ventas-afiliado/", success_ventas_afiliado_view, name="success_ventas_afiliado"),
]
