# Dashboard HH Mantención SAP

**Piscicultura Lago Verde**

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

## Cambios v3.10.19

- Agrega la columna `Operación` después de `N° orden` en el detalle del plan para los archivos desde Semana 37 en adelante.
- Incluye `Operación` al copiar filas seleccionadas y al exportar CSV.
- Mantiene compatibles los planes anteriores que no contienen esa columna.

## Cambios v3.10.18

- Todos los gráficos usan el período seleccionado: semanal, mensual o anual.
- En modo semanal, el gráfico principal compara únicamente la semana elegida contra 87,5 HH.
- Los títulos indican la semana seleccionada en vez de mostrar los meses que atraviesa esa semana.
- En modo mensual se muestra solo el mes consultado; en modo anual se mantiene el desglose mensual.

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
