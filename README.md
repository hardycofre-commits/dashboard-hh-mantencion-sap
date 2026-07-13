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
- Informe PDF ejecutivo de dos páginas.
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
