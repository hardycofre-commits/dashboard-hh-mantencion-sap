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

## Cambios v3.9.2

- Copia con un clic los números de Aviso y Orden del Plan Semanal.
- Destacado suave al pasar el mouse.
- Confirmación breve después de copiar.
- Se mantiene el mismo diseño de la versión estable.



## Generación y envío automático de informes PDF

El flujo `.github/workflows/informe-automatico.yml` se ejecuta cada vez que se agrega o actualiza un Excel dentro de `datos/` en la rama `main`. El proceso abre el dashboard actualizado, genera automáticamente el informe completo que entrega el botón **Informe PDF**, guarda una copia histórica en `informes/` y la envía por correo.

Para activar el envío, agrega estos secretos en **GitHub > Settings > Secrets and variables > Actions**:

- `SMTP_HOST`: servidor SMTP (para Microsoft 365: `smtp.office365.com`).
- `SMTP_PORT`: puerto SMTP (para Microsoft 365: `587`).
- `SMTP_USER`: cuenta que enviará el correo.
- `SMTP_PASSWORD`: contraseña de aplicación de la cuenta remitente.
- `DESTINATARIOS_INFORME`: correo del planificador; posteriormente puede contener varios correos separados por comas.

El flujo también conserva durante 90 días una copia descargable en GitHub Actions. Puede probarse manualmente desde **GitHub > Actions > Generar y enviar informe PDF > Run workflow**.

<!-- GitHub Pages redeploy 2026-08-06 -->
