# Diccionario de datos (CORRECCION_BASE_ACTUAL_limpio.xlsx)

## BASE DE DATOS
- razon_social: Nombre legal de la empresa.
- ruc: RUC en texto (13 digitos, conserva ceros).
- fecha_afiliacion: Fecha de afiliacion.
- ciudad, direccion, telefonos, emails: Contacto general.
- actividad: Actividad economica declarada.
- contactos/gerentes: Datos de responsables (general, finanzas, RRHH, comercial, produccion).

### Ventas
- ventas_2019, ventas_2020, ventas_2021, ventas_2022, ventas_2023: Ventas anuales. Formato moneda #,##0.00. Vacio si no reporta.

### Tamano
- tamano_cod_2022 / tamano_cod_2023: Codigo de tamano por ano (1=MICRO, 2=PEQUENA, 3=MEDIANA, 4=GRANDE).
- tamano_2022 / tamano_2023: Etiqueta de tamano derivada del codigo.
- tamano_actual: Tamano textual consolidado (ultima version conocida).
- cambio_tamano: Comparacion 2022->2023 (MISMO, SUBE, BAJA; NA si falta alguno).
- porcentaje: Porcentaje asociado (si aplica en dashboards).
- semaforo: Estado semaforico del cliente (ROJO/AMARILLO/VERDE).
- estado: Estado de pago (PAGADO/NO PAGADO).

## CAPACITACIONES
- razon_social: Identificador de la empresa.
- q1_capacitaciones ... q4_capacitaciones: Numero de capacitaciones por trimestre.
- q1_valor ... q4_valor: Valor monetario por trimestre (moneda #,##0.00).
- total_capacitaciones: Suma de q1-q4 (conteo).
- valor_total: Suma de q1_valor-q4_valor.

## DIAGNOSTICOS
- razon_social: Identificador de la empresa.
- lean, estrategia, legal, ambiente, rrhh: Indicadores (1 si marcado, 0 si no). Se interpretan X como 1.
- total: Suma de los indicadores.
- diagnostico: SI si total>0, NO en caso contrario.

## LEGAL
- razon_social: Identificador de la empresa.
- contacto: 1 si hubo contacto marcado con X, 0 si no.
- laboral, societario, propiedad_intelectual, otros: Indicadores (sumados desde LEGAL y LEGAL (2), X convertido a 1).
- total: Suma de los indicadores.
- servicio: SI si total>0, NO en caso contrario.

## Notas de limpieza
- Placeholders (-, X, null, pendiente, no reporta, comillas/espacios) se normalizaron a NA donde corresponde.
- RUC se mantiene como texto; importar con dtype={'ruc': str} para no perder ceros.
- Formatos monetarios aplicados en ventas_* y valores de capacitaciones.
- Hoja SCHEMA incluye el diccionario de columnas con tipo y valores permitidos por hoja.
