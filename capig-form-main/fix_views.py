import re

# Leer el archivo corrupto
with open('forms/view/form_views.py', 'r', encoding='utf-8') as f:
    content = f.read()

# Buscar y eliminar código duplicado de diagnóstico dentro de nuevo_afil iado_view
# Encontrar el inicio de nuevo_afiliado_view
pattern = r'(@require_http_methods\[\["GET", "POST"\]\]\s+def nuevo_afiliado_view\(request\):\s+"""Formulario para registrar un nuevo afiliado en la hoja BASE DE DATOS\."""\s+sectores = _obtener_sectores\(\)\s+)"""Vista para el formulario de diagnóstico""".*?return render\(request, \'diag_form\.html\', \{\'empresas\': empresas\}\)\s+'

# Si encontramos el patrón, lo eliminamos
if '"""Vista para el formulario de diagnóstico"""' in content:
    lines = content.split('\n')
   
    cleaned_lines = []
    skip = False
    skip_count = 0
    
    for i, line in enumerate(lines):
        # Detectar inicio de la duplicación
        if 'def nuevo_afiliado_view(request):' in line:
            cleaned_lines.append(line)
            # Añadir las líneas correctas hasta el POST
            if i+1 < len(lines):
                cleaned_lines.append(lines[i+1])  # docstring
            if i+2 < len(lines):
                cleaned_lines.append(lines[i+2])  # sectores
            if i+3 < len(lines):
                cleaned_lines.append(lines[i+3])  # blank line
           
            # Saltar la duplicación (buscar hasta encontrar el verdadero if request.method)
            j = i + 4
            while j < len(lines) and not ('if request.method == "POST":' in lines[j] and 'razon_social = request.POST.get("razon_social"' in lines[j+5] if j+5 < len(lines) else False):
                j += 1
           
            # Copiar el resto desde el verdadero POST
            for k in range(j, len(lines)):
                cleaned_lines.append(lines[k])
            break
        elif not skip:
            cleaned_lines.append(line)
    
    content = '\n'.join(cleaned_lines)

# Guardar el archivo limpio
with open('forms/view/form_views.py', 'w', encoding='utf-8') as f:
   f.write(content)

print("Archivo reparado correctamente")
