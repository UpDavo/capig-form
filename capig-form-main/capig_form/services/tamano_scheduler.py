"""
Ejecucion periodica del job de tamanos sin tocar el resto del sistema.

Por defecto corre cada 15 minutos. Puedes ajustar con la variable de
entorno TAMANO_JOB_INTERVAL_SECONDS (minimo 300 para evitar abusos).

Uso:
    python -m capig_form.services.tamano_scheduler
Interrumpir con Ctrl+C cuando no se necesite.
"""

import os
import time

from capig_form.services import tamano_empresas_job as job


def _get_interval_seconds() -> int:
    """Obtiene el intervalo, con piso de 300s para no bombardear la API."""
    raw = os.getenv("TAMANO_JOB_INTERVAL_SECONDS")
    try:
        seconds = int(raw) if raw else 900  # default 15 min
    except (TypeError, ValueError):
        seconds = 900
    return max(seconds, 300)


def main():
    interval = _get_interval_seconds()
    print(f"[tamano_scheduler] Iniciando loop cada {interval} segundos. Ctrl+C para salir.")
    while True:
        start = time.time()
        try:
            job.run()
            print("[tamano_scheduler] Ejecucion completada.")
        except Exception as exc:  # pragma: no cover - modo servicio
            # No interrumpimos el loop; solo informamos.
            print(f"[tamano_scheduler] Error: {exc}")
        elapsed = time.time() - start
        sleep_for = max(interval - elapsed, 0)
        time.sleep(sleep_for)


if __name__ == "__main__":
    main()
