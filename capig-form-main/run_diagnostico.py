import os
import sys

import django

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

os.environ.setdefault("DJANGO_SETTINGS_MODULE", "capig_form.settings")
django.setup()

from capig_form.services.diagnostico_generator import generar_diagnostico_y_subir

if __name__ == "__main__":
    generar_diagnostico_y_subir()
    print("Diagnostico generado y subido correctamente.")
