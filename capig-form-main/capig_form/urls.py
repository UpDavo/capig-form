from django.contrib import admin
from django.urls import path, include
from forms.view.form_views import home_view

urlpatterns = [
    path('', home_view, name='home'),              # Página de inicio
    path('admin/', admin.site.urls),
    path('', include('forms.urls')),               # Rutas de los formularios
]

handler404 = 'forms.view.form_views.custom_404_view'
