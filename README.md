# Dashboard HH Mantención SAP

**Piscicultura Lago Verde**

## Línea de tiempo de datos (v3.11.4)

- **Hasta la Semana 37 de 2026:** se conserva el sistema histórico actual, incluidos los Excel `SemanaXX.xlsx`, las HH netas históricas y los estados de terreno consolidados mediante Google Sheets.
- **Desde la Semana 38 de 2026:** se usa el nuevo export SAP. La fecha de corte es el lunes **14-09-2026** y está centralizada en `FECHA_CORTE_SEMANA_38_2026`.
- Cuando el export anual trae `Fecha de fin de ejecución real`, este archivo pasa a ser la fuente completa del resumen de HH. El corte del 14-09-2026 se mantiene para separar el Plan Semanal histórico del nuevo.
- Las consultas Semana, Mes y Año combinan ambas fuentes cuando el período cruza el corte.

## Nuevo export SAP

El archivo se detecta automáticamente entre los Excel de `datos/`. Sus encabezados se reconocen por nombre normalizado y no por posición. Columnas requeridas:

- `Inic.extr.` o `Fecha de inicio extrema`: fecha principal de planificación.
- `Fecha de fin de ejecución real`: fecha usada para imputar las HH y las órdenes ejecutadas al resumen semanal, mensual o anual.
- `Aviso` y `Orden`.
- `Texto breve`: descripción breve asociada a la orden.
- `Op.` u `Operación`.
- `Texto breve operación`: descripción principal del trabajo.
- `Pto.tbjo.op.` o `Pto.tbjo.operación`: puesto de trabajo conservado en el modelo. Para esta etapa corresponde al responsable único **Asistente de mantención**.
- `Status sistema op.`: única fuente del estado nuevo.
- `Trabajo real`: HH reales directas; por ejemplo, `1,500` se interpreta como `1.5`.

Desde el corte, cada registro se identifica por **Orden + Operación**. No se elimina una operación por tener cero HH y no se descartan fechas futuras.

## Cambios v3.11.1

- Se incorpora `Texto breve` de la orden en el detalle del Plan Semanal.
- El texto breve de la orden también se incluye al copiar filas, exportar CSV y generar el detalle PDF.

## Cambios v3.11.2

- Al abrir Plan Semanal, el filtro cambia automáticamente a la semana ISO vigente; avanzará a las semanas 39, 40 y siguientes según la fecha del sistema.
- Se centra y estabiliza el ancho de la columna Estado para alinear el encabezado con sus etiquetas.

## Cambios v3.11.3

- Se elimina del filtro superior la etiqueta redundante `META HH MENSUAL`; la meta sigue funcionando internamente y se mantiene visible en los indicadores del dashboard.

## Cambios v3.11.4

- Se separa la fecha de planificación de la fecha real de finalización.
- El Plan Semanal, los estados vencidos y los estados notificados continúan controlándose desde la fecha planificada y el status SAP.
- Las HH, las órdenes ejecutadas, el resumen diario y los gráficos se contabilizan en la `Fecha de fin de ejecución real`.
- Las operaciones sin fecha real de finalización no se suman al resumen de ejecución.

## Estados desde Semana 38

El estado se normaliza y evalúa en este orden:

1. Si contiene `INBO`, la operación se excluye completamente.
2. Si comienza con `CTEC` o `CETEC`, se excluye. `NOTI CTEC...` no se excluye porque CTEC no está al comienzo.
3. Si contiene `NOTI`, queda **Notificada**.
4. Si contiene `LIB.`, queda **Vencida**, **Para hoy** o **Programada** según la fecha frente al día actual.

Las semanas se calculan en formato ISO, de lunes a domingo, desde la Fecha de inicio extrema. Las operaciones Programadas permanecen visibles, pero no integran el denominador del cumplimiento del Plan hasta que su fecha sea evaluable. Si toda la consulta es futura, el cumplimiento se muestra como `—`.

El cumplimiento HH conserva las metas operacionales: 11,7 HH por día, 87,5 HH por semana y 350 HH por mes; una consulta anual suma 350 HH por cada mes incluido.

Encargado, Turno, sus filtros y el gráfico por encargado se conservan solamente para consultas con información histórica. Google Sheets no determina ningún estado desde la Semana 38.

## Estructura del proyecto

- `index.html`: estructura principal del dashboard.
- `styles.css`: estilos visuales y formato de impresión/PDF.
- `app.js`: lectura de archivos SAP y Plan Semanal, cálculos, gráficos y exportación PDF.
- `README.md`: descripción y guía breve del proyecto.

## Datos

El dashboard consulta automáticamente los archivos Excel almacenados en la carpeta:

```text
datos/
```

Características actuales:

- Selección automática del último archivo SAP cargado en GitHub.
- Selector histórico de Plan Semanal.
- Cálculo de HH netas por orden.
- Tratamiento de anulaciones SAP.
- Resumen, gráficos y detalle del Plan Semanal.
- Informe PDF en formato Carta vertical con selector de secciones: resumen ejecutivo, resumen diario y plan semanal.
- Selección automática del período en modo mensual o anual, meta base fija de 350 HH por mes y actualización inmediata al cambiar filtros.
- Paginación automática del resumen diario en informes anuales.
- Tablas seleccionables para copiar Aviso u Orden con `Ctrl + C`.

## Publicación en GitHub Pages

Los archivos deben quedar en la raíz del repositorio:

```text
index.html
app.js
styles.css
README.md
datos/
```

Después de reemplazar archivos, realiza un commit y actualiza el sitio con `Ctrl + F5`.

## Cambios v3.10.17

- El modo semanal usa una meta total fija de 87,5 HH.
- La meta diaria se mantiene en 11,7 HH, calculada como 350 HH mensuales dividido por 30 días.
- Los modos mensual y anual conservan su cálculo anterior.

## Cambios v3.10.16

- Corrige el identificador de caché del script principal, que todavía apuntaba a v3.10.13.
- Fuerza una lectura fresca del último Excel para evitar reutilizar exports guardados por el navegador.
- Verificado con el export del 7 de septiembre: 968 registros totales y 4,0 HH correspondientes a septiembre.

## Cambios v3.10.15

- Abre automáticamente el mes calendario en curso al iniciar el dashboard.
- Usa un índice alternativo para encontrar el último export SAP cuando la API pública de GitHub está limitada.
- Actualiza el respaldo hasta el export del 7 de septiembre de 2026 para que las HH de septiembre se carguen inmediatamente.

## Cambios v3.10.14

- Corrige la ruta de `Semana28.xlsx` respetando mayúsculas y minúsculas en GitHub Pages.
- Incorpora `Semana35.xlsx` y `Semana36.xlsx` a la lista de respaldo.
- Detecta automáticamente nuevos archivos `SemanaNN.xlsx` publicados, incluso cuando la API pública de GitHub está temporalmente limitada.

## Cambios v3.9.2

- Copia con un clic los números de Aviso y Orden del Plan Semanal.
- Destacado suave al pasar el mouse.
- Confirmación breve después de copiar.
- Se mantiene el mismo diseño de la versión estable.



<!-- GitHub Pages redeploy 2026-08-06 -->
