# Generador de Memos de Laboratorio

PWA (página web sin instalación) que genera los Excel de envío de muestras al
Departamento de Laboratorio de SERNAGEOMIN, con el **formato institucional exacto**
(membrete y logos incluidos), a partir de la tabla de muestras del proyecto.

Pensada para uso cotidiano de todo el equipo: se abre en cualquier navegador, no
requiere ArcGIS ni instalar nada.

## Qué hace

1. Pegas la tabla de muestras (copiada desde la **tabla de atributos de ArcGIS**,
   Ctrl+C) o arrastras un **CSV**.
2. La app lista las muestras, detecta el tipo de análisis por la sigla
   (`-T` cortes transparentes · `-D` datación · `-x` rayos X) y te deja
   seleccionar cuáles enviar.
3. Con un clic genera **un solo archivo `.xlsx`** que contiene:
   - **Ficha de Ingreso** con todas las muestras seleccionadas.
   - Un **Memo** por cada tipo presente (Cortes / Datación U-Pb / Rayos X).
   - Una **Ficha de Datación** por cada muestra de datación.

El **número de memorando y la firma se completan a mano en Excel** (no se
automatizan a propósito, por ser un documento formal).

## Cómo funciona por dentro (resumen técnico)

La plantilla `assets/plantilla_base.xlsx` es un memo real "en blanco" con
**tokens** `@@...@@` sembrados en cada celda rellenable. En el navegador la app
descomprime el `.xlsx` (es un ZIP), **reemplaza los tokens como texto** en el XML
de las hojas y vuelve a comprimir — **sin tocar los logos/membrete** (que viven en
`xl/media` + `xl/drawings`), así el formato institucional queda intacto byte a
byte. No se usa ninguna librería que tenga que "entender" el Excel: solo JSZip
para el ZIP y reemplazo de texto.

Las hojas de tipos no usados y las fichas de datación sobrantes se marcan como
ocultas en el archivo final.

## Estructura

```
src/app.jsx            aplicación React (parser, tabla de selección, generación)
src/template.html      HTML base con placeholders para el build single-file
assets/plantilla_base.xlsx   plantilla con tokens (generada por scripts/)
scripts/construir_plantilla_base.py   regenera la plantilla desde un memo real
vendor/                React, ReactDOM, Babel, JSZip (todo local, sin CDN)
build.ps1              genera dist/GeneradorMemos_vN.html + pwa/index.html
serve.ps1              sirve pwa/ en http://localhost:8130 para probar/instalar
make-icons.ps1         genera los iconos de la PWA
pwa/                   PWA lista para GitHub Pages (index.html, sw.js, manifest, icons)
dev.html               versión de desarrollo (carga vendor/ y src/ sin compilar)
docs/SPEC.md           documento de diseño
```

## Desarrollo

```powershell
# Probar la versión de desarrollo (sin compilar): servir la carpeta y abrir dev.html
python -m http.server 8130 ; # luego abrir http://localhost:8130/dev.html

# Regenerar la plantilla desde un memo real (solo si cambia el formato)
python scripts/construir_plantilla_base.py "ruta\a\Memo N°XX.xlsx"

# Compilar el single-file + la PWA
./build.ps1

# Servir la PWA para probar la instalación / offline
./serve.ps1   # http://localhost:8130
```

## Publicar en GitHub Pages

Se publica la carpeta `pwa/` (contiene `index.html` autocontenido + `sw.js` +
`manifest` + `icons`). Ver la memoria del proyecto para el flujo de publicación.

## Límites de la versión actual

- Ficha de Ingreso: hasta 24 muestras · Memo por tipo: hasta 12 · Fichas de
  datación: hasta 3 por archivo (según el espacio de las plantillas reales). Si
  se excede, la app avisa y pide dividir el envío.
- La lectura directa de una `.gdb`/shapefile en el navegador (arrastrar la
  geodatabase) queda como mejora futura; hoy la entrada es copiar/pegar o CSV.
- El campo "unidad geológica" se toma de la columna correspondiente o, si no
  existe, de OBSERVACIONES, y es **editable** en la tabla antes de generar.
