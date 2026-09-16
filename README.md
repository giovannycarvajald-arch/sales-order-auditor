# Sales Order Auditor — V1

Auditor local para Sales Orders de Odessa Separator Inc.

## Objetivo

Cargar un PDF de una Sales Order, extraer los datos principales y ejecutar reglas determinísticas de auditoría.

## Privacidad

- No hay backend.
- No hay API key.
- El PDF se procesa dentro del navegador.
- Price List, borradores e historial se guardan en el almacenamiento local del navegador.
- GitHub Pages solo sirve el código estático.

> Nota: esta V1 carga PDF.js y SheetJS desde CDN. El contenido del PDF no se envía a esos servicios; los scripts son dependencias de la aplicación. En una V2 podemos incluir las librerías dentro del repositorio para eliminar esa dependencia externa.

## V1 incluye

- PDF → extracción de texto en navegador.
- Campos editables para corregir extracción.
- Price List XLSX local.
- Tax 8.25% para clientes clasificados como taxable.
- Coterra/Diamondback → Tax $0.
- BTA → taxable.
- XTO → STAMP.
- Diamondback → AFE & GL.
- Delivery: END USER / AFTER HOURS / Pickup / Summit.
- Inspection Fee 1004-0009-00 cuando aparece 1014-0142-00.
- Comparación de precios contra Price List.
- Validación de teléfono.
- Sold To / Ship To y well.
- Observación GRS solo cuando la configuración lo sugiere y no hay GRS.
- Historial local.

## GitHub Pages

1. Crea un repositorio privado en GitHub.
2. Sube `index.html`, `styles.css`, `app.js` y `README.md`.
3. En Settings → Pages, configura GitHub Actions como fuente.
4. Abre la URL de GitHub Pages.

Para máxima privacidad del código, usa repositorio privado. La aplicación sigue siendo accesible para ti mediante GitHub Pages según la configuración de tu cuenta/organización.

## Próximas mejoras

- Parser específico para el formato exacto de las SO de Odessa.
- Importación robusta de Price List por encabezados.
- Base local de PADs.
- Regla "solo primera SO del PAD cobra Delivery".
- Catálogo editable de clientes.
- Reporte imprimible/PDF.
- Exportación a Excel.
- Tests automáticos con SO ya auditadas.
- V2 opcional con IA para casos ambiguos, manteniendo cálculos críticos determinísticos.

### V1.1 parser fix
The item parser now attempts to read unit prices from the PDF text instead of creating detected part numbers with a default price of $0.00. Because PDF layouts can vary, the extracted item values remain editable before auditing.

### V1.2 parser
Improved PDF line reconstruction, Odessa item-row parsing, stacked Total/Tax/Subtotal parsing, and phone/contact extraction.

### V1.3 parser fix
The item parser supports Odessa rows with or without a populated Rev column and descriptions wrapped across multiple PDF lines. Each item row is parsed independently.

### V1.4 parser fix
The parser now finds every unique part number inside the Odessa item-table area and parses each item independently from the quantity/unit-price/amount ending. This is designed to handle wrapped descriptions and inconsistent Rev-column population.
