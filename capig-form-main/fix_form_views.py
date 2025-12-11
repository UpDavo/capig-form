# Script para agregar el campo colaboradores a form_views.py
import re

# Leer archivo
with open('forms/view/form_views.py', 'r', encoding='utf-8', errors='ignore') as f:
    lines = f.readlines()

# Buscar la función nuevo_afiliado_view
start_idx = None
for i, line in enumerate(lines):
    if 'def nuevo_afiliado_view(request):' in line:
        start_idx = i
        break

if start_idx is None:
    print("ERROR: No se encontró la función nuevo_afiliado_view")
    exit(1)

# Eliminar contenido corrupto desde start_idx + 4 hasta encontrar el verdadero if POST
real_post_idx = None
for i in range(start_idx + 4, len(lines)):
    # Buscar la línea que tiene GET/POST y luego razon_social
    if 'if request.method == "POST":' in lines[i]:
        # Verificar que las siguientes líneas contienen el código de afiliado (not diagnóstico)
        next_lines = ''.join(lines[i:min(i+10, len(lines))])
        if 'razon_social = request.POST.get("razon_social"' in next_lines:
            real_post_idx = i
            break

if real_post_idx is None:
    print("ERROR: No se encontró el POST real")
    exit(1)

# Eliminar código corrupto (docstring duplicado + código de diagnóstico)
new_lines = lines[:start_idx+4] + lines[real_post_idx:]

# Ahora agregar colaboradores
modified = False
for i, line in enumerate(new_lines):
    if 'genero = request.POST.get("genero", "").strip()' in line:
        # Insertar colaboradores después de genero
        indent = line[:len(line) - len(line.lstrip())]
        new_lines.insert(i+1, f'{indent}colaboradores' + ' = request.POST.get("colaboradores", "").strip()\n')
        modified = True
        break

if not modified:
    print("ERROR: No se pudo agregar colaboradores")
    exit(1)

#  Agregar colaboradores en el diccionario pasado a guardar_nuevo_afiliado_en_google_sheets
for i, line in enumerate(new_lines):
    if '"genero": genero,' in line:
        indent = line[:len(line) - len(line.lstrip())]
        new_lines.insert(i+1, f'{indent}"colaboradores": colaboradores,\n')
        break

# Escribir archivo
with open('forms/view/form_views.py', 'w', encoding='utf-8', newline='') as f:
    f.writelines(new_lines)

print("✅ Archivo form_views.py reparado y actualizado correctamente")
