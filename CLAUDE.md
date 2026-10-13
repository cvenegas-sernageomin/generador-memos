# Generador de memos de laboratorio — notas para Claude

PWA: tabla de muestras (TSV pegado o CSV) → un .xlsx con el membrete institucional exacto.
Repo `cvenegas-sernageomin/generador-memos` (rama `main`). Ficha en memoria: `proyecto-generador-memos-pwa`.

## Reglas
- La plantilla `assets/plantilla_base.xlsx` lleva tokens `@@...@@`; en el navegador se reemplazan como texto
  en `xl/worksheets/sheetN.xml` con JSZip **sin tocar** `xl/media`/`xl/drawings` (exceljs no abre estas plantillas).
  Ver memoria `reference-excel-membrete-tokens-jszip-navegador`.
- `src/` es la fuente; `build.ps1` genera `dist/` (single-file con la plantilla en base64) y `pwa/index.html`. Subir `CACHE` en `pwa/sw.js` al cambiar la PWA.
- Número de memorando y firma quedan **manuales** (documento formal).

## Verificar
Abrir el .xlsx generado con openpyxl: sin reparación, 20 logos intactos, valores correctos.
