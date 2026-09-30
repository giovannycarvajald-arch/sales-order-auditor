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

> Nota: esta V1 carga PDF.js y SheetJS desde CDN. El contenido del PDF no se envía a esos servicios; los scripts son dependencias de la aplicación. En una V3 podemos incluir las librerías dentro del repositorio para eliminar esa dependencia externa.

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
- V3 opcional con IA para casos ambiguos, manteniendo cálculos críticos determinísticos.

### V1.1 parser fix
The item parser now attempts to read unit prices from the PDF text instead of creating detected part numbers with a default price of $0.00. Because PDF layouts can vary, the extracted item values remain editable before auditing.

### V1.2 parser
Improved PDF line reconstruction, Odessa item-row parsing, stacked Total/Tax/Subtotal parsing, and phone/contact extraction.

### V1.3 parser fix
The item parser supports Odessa rows with or without a populated Rev column and descriptions wrapped across multiple PDF lines. Each item row is parsed independently.

### V1.4 parser fix
The parser now finds every unique part number inside the Odessa item-table area and parses each item independently from the quantity/unit-price/amount ending. This is designed to handle wrapped descriptions and inconsistent Rev-column population.

### V1.5 parser fix
The parser now works from the PDF's visual text lines. Each Part Number starts an independent block, wrapped descriptions are joined, and the item is accepted as soon as a `Qty UnitPrice/EA Amount` pattern is found. This avoids the `Ship Dates` text from swallowing the next item's price.

### V1.6 parser fix
The item detector no longer requires the line number and Part Number to be in the same PDF.js visual line. It finds every unique Part Number in the item-table area and parses the first quantity/unit-price/amount pattern after that PN. This addresses PDFs where PDF.js separates table columns into different text runs.

### V1.7 parser fix
The item area no longer ends at "Return Policy". In Odessa multi-page PDFs, Return Policy is printed before the actual item rows, so the previous boundary produced zero detected items. V1.7 scans from the first item-table header through the rest of the document and deduplicates Part Numbers.

### V1.8 UI/file-load fix
The interface now clearly separates "PDF loaded" from "items detected". Selecting or dropping a PDF immediately confirms the filename and size. Clicking Analyze shows processing status and reports separately whether the PDF was successfully read and whether item rows were found.

### V1.9 Price List fix
The importer now uses the exact Odessa workbook structure: `PN LIST` → `PartNumber`, `Name`, `Pricing_UnitPrice0`. It no longer guesses the first numeric value in a row. Price lookups are exact normalized Part Number matches, and the loaded Price List is persisted locally in the browser.

### V3.2 Procedure-driven audit + Price List correction
Price List loading now strictly uses `PN LIST` and the `Pricing_UnitPrice0` column. It no longer guesses numeric cells. The app uses a versioned local-storage key so stale mappings from earlier versions cannot override the newly loaded Price List. Missing PNs are explicitly reported as `PN NOT FOUND IN PRICE LIST`.


### V3.2 changes
- Added a local PDF uploader for the current Sales Order Procedure.
- The app reads the section **7. INFORMATION TO CONSIDER WHEN CREATING THE SALES ORDER** and extracts the `TAXABLE`, `STAMP`, and `DISCOUNT` lists from the procedure (the current REV30 lists are on page 22).
- Procedure rules are stored in browser `localStorage` and are replaced when a newer procedure is uploaded.
- Tax, STAMP, and discount checks use the loaded procedure instead of the old hardcoded customer lists.
- Added a `Disc%` column to SO items so procedure discounts can be verified line-by-line. Delivery and inspection charges are excluded from the customer discount check.
- Price List loading is strict: sheet `PN LIST`, columns `PartNumber` and `Pricing_UnitPrice0`, stored under versioned `soPriceList_v2` local storage.


### Exxon Mobil → XTO alias rule
- The Procedure's TAXABLE and STAMP lists contain `EXXON MOBIL (ALL EXXON ORDERS WILL BE XTO)`.
- The auditor now explicitly interprets this as an alias rule: any Sales Order whose customer is `XTO`, `XTO ENERGY`, or `XTO ENERGY INC.` inherits the Exxon Mobil TAXABLE and STAMP requirements.
- The application does not require the SO itself to contain the word Exxon Mobil.
- The audit explanation identifies the rule as `EXXON MOBIL → XTO` so the result is traceable to the Procedure.

### V3.5 Procedure alias fix
REV30 aliases are now structured during Procedure loading. The explicit rule `EXXON MOBIL (ALL EXXON ORDERS WILL BE XTO)` is stored as `EXXON MOBIL -> XTO` and is used for both TAXABLE and STAMP. The audit no longer hardcodes XTO as taxable/stamp; it derives that relationship from the loaded Procedure. `index.html` cache-busts `app.js?v=2.5` to prevent GitHub Pages/browser caching of an older auditor script.

### V3.6 Procedure customer matching
XTO ENERGY / XTO ENERGY INC are normalized to the same customer key `XTO` for dynamic Procedure matching. If the loaded Procedure lists XTO ENERGY or EXXON MOBIL, TAXABLE and STAMP rules are applied to XTO ENERGY INC on the Sales Order. Procedure rules are stored under v2 keys to prevent stale cached rules from previous versions.

### V3.7 Procedure parser fix
Procedure sections are now parsed independently by locating TAXABLE, STAMP, DISCOUNT, and CONFIRMATION CHECK. XTO ENERGY entries are retained explicitly, and XTO/XTO ENERGY/XTO ENERGY INC. are treated as the same customer for procedure matching. Procedure local storage was versioned to prevent stale rules from earlier builds.

### V3.8 Procedure parser fix
Procedure sections are parsed independently by TAXABLE/STAMP/DISCOUNT/CONFIRMATION CHECK boundaries. XTO ENERGY is canonicalized so XTO, XTO ENERGY and XTO ENERGY INC on a Sales Order match the Procedure. The Procedure load status now explicitly reports whether XTO was found in TAXABLE and STAMP.


### V3.2 Delivery extraction
The SO parser now prioritizes the explicit Odessa note format `Delivery: M/D/YYYY h:mm AM/PM` when extracting Delivery Date and Delivery Time. The detected line is also shown in Delivery Text as source context.


### V3.4 — Well extraction
The PDF parser now treats the line immediately below the company in the Sold To / Ship To party block as a potential well name. It recognizes well names such as `NAIL TILLMAN W1P 416H` and `JORDAN 23-24 2814H`, while excluding address/header lines. Cache-busting was also updated to `app.js?v=3.4`.

### V3.5 — Ship To Well extraction
Odessa Sales Orders place the well on the line immediately below the company in the Ship To block. The parser now reads rows below the Ship To header, keeps Sold To billing-address rows out of well detection, and does not infer a well from Sold To.
