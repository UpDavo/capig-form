from django.http import JsonResponse
from django.views.decorators.http import require_http_methods
from django.views.decorators.csrf import csrf_exempt


@csrf_exempt
@require_http_methods(["GET", "HEAD"])
def health_check(request):
    """
    Health check endpoint para Railway.
    Retorna 200 OK si la aplicación está funcionando correctamente.
    """
    return JsonResponse({"status": "healthy"}, status=200)
