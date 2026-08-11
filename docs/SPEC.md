# Generador de Memos de Laboratorio — Diseño (as-built)

**Fecha:** 2026-08-11 · **Estado:** v1 construida y verificada de punta a punta.

## 1. Objetivo

PWA sin instalación que genera los Excel de envío de muestras al laboratorio con el
formato institucional exacto (membrete/logos), para uso cotidiano de todo el equipo.
Reemplaza el copiar-editar-a-mano el último memo.

## 2. Entrada de datos

Copiar/pegar la selección de la **tabla de atributos de ArcGIS** (TSV) o arrastrar un
**CSV**. Como todo el equipo usa la misma gdb, las columnas son consistentes. Columnas
reconocidas por alias (mayúsculas sin acentos): `SIGLA_MUESTRA`, `TIPO_ANALISIS`,
`LITOLOGIA`, `UNIDAD_GEOLOGICA`/`OBSERVACIONES`, `GEOLOGO`, `COORD_ESTE`, `COORD_NORTE`,
`TIPO_MUESTRA`, `FECHA_MUESTREO`. Solo `SIGLA_MUESTRA` es obligatoria.

> Descartado para v1: leer la `.gdb`/shapefile directamente en el navegador con
> gdal3.js. Motivo: no verificable sin la gdb (bloqueada por sincronización OneDrive
> durante el desarrollo), ~40 MB de WASM, y el usuario priorizó "lo más sencillo".
> Copiar/pegar cubre el caso y es 100% testeable. Queda como mejora futura.

## 3. Detección de tipo

Por sufijo de la sigla tras el último dígito: última letra `t`→cortes, `d`→datación,
`x`→rayos X (ej. `CV251008-2T`, `C281125-6Ad`, `CV260430-1x`). Fallback: `TIPO_ANALISIS`.
Editable por muestra en la tabla; si queda `?` la generación se bloquea con aviso.

## 4. Generación del Excel — enfoque de tokens (clave del diseño)

`assets/plantilla_base.xlsx` es un memo real "en blanco" con **tokens únicos**
`@@...@@` sembrados en cada celda rellenable (openpyxl los escribe como `inlineStr`).
En el navegador: JSZip descomprime el `.xlsx`, se **reemplazan los tokens como texto**
(con escape XML) en `xl/worksheets/sheetN.xml`, y se recomprime. Los tokens sin valor
se blanquean.

**Por qué así y no exceljs:** exceljs (y cualquier lib que reserialice el Excel) **no
puede ni abrir** estas plantillas — falla en el parseo de los `drawings` del membrete
(`Cannot read properties of undefined (reading 'anchors')`). El enfoque de tokens no
toca `xl/media` ni `xl/drawings`, así que los logos se preservan byte a byte.
Verificado: el archivo generado conserva las 20 imágenes de logos y abre sin reparación.

Hojas de tipos ausentes y fichas de datación sobrantes se marcan `state="hidden"` en
`xl/workbook.xml`.

## 5. Contenido del archivo generado (un solo `.xlsx`)

- **DatosProyecto** (valores por defecto del proyecto, fijos).
- **Ficha de Ingreso**: una fila por muestra seleccionada (ítem, requerimiento, sigla,
  proyecto, centro de costo, colector, UTM, tipo, litología, unidad, fecha, obs).
- **Memo N° CT / U-Pb / RX**: uno por cada tipo presente; fecha + listado (número/sigla,
  hasta 12 en dos pares de columnas). MAT y párrafo de solicitud vienen fijos de la
  plantilla real. **Número de memorando y firma: manuales.**
- **Ficha Datación N**: una por muestra de datación (solicitante, id, litología, unidad,
  clasificación); edad/descripciones/observaciones quedan en blanco (juicio del geólogo).

## 6. Límites (según el espacio de las plantillas reales)

Ficha de Ingreso ≤ 24 · Memo por tipo ≤ 12 · Fichas de datación ≤ 3. Si se excede, la
app avisa y pide dividir el envío.

## 7. Empaquetado

React + Babel + JSZip vendorizados (sin CDN, sin Node). `build.ps1` embebe la plantilla
como base64 y genera un `dist/GeneradorMemos_vN.html` autocontenido (doble clic, offline)
y `pwa/index.html` para GitHub Pages (con `sw.js` para uso offline).

## 8. Verificación realizada

Flujo completo en navegador (paste → selección → generar) con muestras de tipos mixtos.
El `.xlsx` real generado, extraído del navegador y abierto con openpyxl: abre sin
reparación, 20 logos intactos, valores correctos en Ficha de Ingreso / memos / fichas,
fecha en castellano, escape XML de caracteres especiales (`& < >`), y ocultamiento
correcto de hojas no usadas. El build single-file usa la plantilla embebida (sin fetch).
