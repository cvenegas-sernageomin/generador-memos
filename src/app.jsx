const { useState, useMemo, useRef } = React;

/* ================== CONFIG / MAPEO ==================
   Constantes del proyecto (misma gdb para todo el equipo) y alias de columnas.
   Si cambia el proyecto o los nombres de campo, se ajusta acá. */
const CONST = {
  proyecto: "Cuadrángulo Central Los Cipreses",
  centroCosto: "8003",
  tipoMuestra: "Roca",
};

// Texto de "Requerimiento Analítico" que va en la Ficha de Ingreso por tipo detectado.
const REQ_TEXT = { CT: "cortes transparente", DAT: "Datación U/Pb", RX: "rayos x" };
const TIPO_LABEL = { CT: "Cortes transp.", DAT: "Datación", RX: "Rayos X", "?": "(sin tipo)" };

// Alias de nombres de columna aceptados (normalizados: MAYÚSCULAS sin acentos ni separadores).
const CAMPOS = {
  sigla:      ["SIGLAMUESTRA", "IDMUESTRASIGLA", "SIGLA", "ID"],
  tipoAnal:   ["TIPOANALISIS", "REQUERIMIENTO", "REQUERIMIENTOANALITICO"],
  litologia:  ["LITOLOGIA", "LITOLOGIANOMBREROCA", "ROCA"],
  unidad:     ["UNIDADGEOLOGICA", "UNIDAD"],
  observ:     ["OBSERVACIONES", "OBS"],
  geologo:    ["GEOLOGO", "COLECTOR", "COLECTORGEOLOGO"],
  este:       ["COORDESTE", "UTMESTE", "ESTE", "X"],
  norte:      ["COORDNORTE", "UTMNORTE", "NORTE", "Y"],
  tipoMuestra:["TIPOMUESTRA", "TIPODEMUESTRA"],
  fecha:      ["FECHAMUESTREO", "FECHAINGRESOINFORMACION", "FECHA"],
};

// Límites (según formato de las plantillas reales).
const MAX_FI = 24;         // filas de la Ficha de Ingreso
const MAX_LISTADO = 12;    // muestras por hoja Memo (2 pares de columnas × 6 filas)
const POOL_DAT = 3;        // hojas Ficha Datación disponibles en la plantilla

/* ================== UTILIDADES ================== */
function normHeader(h) {
  return (h || "").toString().normalize("NFD").replace(/[̀-ͯ]/g, "")
    .replace(/[^A-Za-z0-9]/g, "").toUpperCase();
}
function escXml(s) {
  return String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;")
    .replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}
const MESES = ["enero","febrero","marzo","abril","mayo","junio","julio","agosto",
  "septiembre","octubre","noviembre","diciembre"];
function fechaMemoLarga(d) { return `Santiago, ${d.getDate()} de ${MESES[d.getMonth()]} de ${d.getFullYear()}`; }
function fechaCorta(d) {
  const p = (n) => String(n).padStart(2, "0");
  return `${p(d.getDate())}/${p(d.getMonth() + 1)}/${d.getFullYear()}`;
}
function detectTipo(sigla, tipoAnal) {
  const s = (sigla || "").trim();
  const m = s.match(/\d([A-Za-z]+)$/);         // letras finales tras el último dígito
  if (m) {
    const last = m[1].slice(-1).toLowerCase();
    if (last === "t") return "CT";
    if (last === "d") return "DAT";
    if (last === "x") return "RX";
  }
  const t = (tipoAnal || "").toLowerCase();
  if (t.includes("corte") || t === "t") return "CT";
  if (t.includes("data") || t.includes("u-pb") || t.includes("u/pb") || t === "d") return "DAT";
  if (t.includes("rayos") || t.includes("rx") || t === "x") return "RX";
  return "?";
}

// Parsea texto pegado (TSV de la tabla de atributos) o CSV. Devuelve {rows, headersRaw}.
function parseTabla(texto) {
  const lineas = texto.replace(/\r/g, "").split("\n").filter((l) => l.trim() !== "");
  if (lineas.length < 2) return { rows: [], headersRaw: [], error: "Pega al menos el encabezado y una fila." };
  const delim = lineas[0].includes("\t") ? "\t" : (lineas[0].includes(";") ? ";" : ",");
  const headersRaw = lineas[0].split(delim).map((h) => h.trim());
  const headersNorm = headersRaw.map(normHeader);
  const idxDe = (claves) => {
    for (const k of claves) { const i = headersNorm.indexOf(k); if (i >= 0) return i; }
    return -1;
  };
  const idx = {};
  for (const campo in CAMPOS) idx[campo] = idxDe(CAMPOS[campo]);
  if (idx.sigla < 0) return { rows: [], headersRaw, error: "No se encontró la columna de sigla de muestra (SIGLA_MUESTRA)." };

  const rows = [];
  for (let i = 1; i < lineas.length; i++) {
    const celdas = lineas[i].split(delim);
    const get = (c) => (idx[c] >= 0 ? (celdas[idx[c]] || "").trim() : "");
    const sigla = get("sigla");
    if (!sigla) continue;
    const unidad = get("unidad") || get("observ");   // unidad geológica: campo propio o, si no, OBSERVACIONES
    rows.push({
      sigla, litologia: get("litologia"), unidad,
      observ: get("observ"), geologo: get("geologo"),
      este: get("este"), norte: get("norte"),
      tipoMuestra: get("tipoMuestra") || CONST.tipoMuestra,
      fecha: get("fecha"),
      tipo: detectTipo(sigla, get("tipoAnal")),
      sel: true,
    });
  }
  return { rows, headersRaw, error: rows.length ? null : "No se leyeron filas con sigla válida." };
}

/* ================== GENERACIÓN DEL XLSX ================== */
const TEMPLATE_B64 = "/*__TEMPLATE_B64__*/";
function b64ToBytes(b64) {
  const bin = atob(b64); const arr = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
  return arr;
}
async function cargarPlantilla() {
  if (TEMPLATE_B64 && !TEMPLATE_B64.startsWith("/*")) return b64ToBytes(TEMPLATE_B64);
  const r = await fetch("assets/plantilla_base.xlsx");
  if (!r.ok) throw new Error("No se pudo cargar la plantilla (assets/plantilla_base.xlsx).");
  return new Uint8Array(await r.arrayBuffer());
}

function construirReemplazos(seleccion) {
  const rep = {};
  const hoy = new Date();
  const fMemo = fechaMemoLarga(hoy), fElab = fechaCorta(hoy);

  // Ficha de Ingreso: todas las muestras seleccionadas
  seleccion.forEach((m, k) => {
    const i = k + 1;
    rep[`FI_item_${i}`] = i;
    rep[`FI_requerimiento_${i}`] = REQ_TEXT[m.tipo] || "";
    rep[`FI_id_${i}`] = m.sigla;
    rep[`FI_proyecto_${i}`] = CONST.proyecto;
    rep[`FI_centro_${i}`] = CONST.centroCosto;
    rep[`FI_colector_${i}`] = m.geologo;
    rep[`FI_este_${i}`] = m.este;
    rep[`FI_norte_${i}`] = m.norte;
    rep[`FI_tipomuestra_${i}`] = m.tipoMuestra;
    rep[`FI_litologia_${i}`] = m.litologia;
    rep[`FI_unidad_${i}`] = m.unidad;
    rep[`FI_fecha_${i}`] = m.fecha;
    rep[`FI_obs_${i}`] = m.observ;
  });

  // Hojas Memo por tipo (listado)
  const porTipo = { CT: [], DAT: [], RX: [] };
  seleccion.forEach((m) => { if (porTipo[m.tipo]) porTipo[m.tipo].push(m); });
  const codeDe = { CT: "CT", DAT: "UPB", RX: "RX" };
  for (const tipo of ["CT", "DAT", "RX"]) {
    const code = codeDe[tipo];
    rep[`${code}_FECHA`] = fMemo;
    rep[`${code}_FELAB`] = fElab;
    porTipo[tipo].forEach((m, k) => {
      rep[`${code}_L${k + 1}_N`] = k + 1;
      rep[`${code}_L${k + 1}_ID`] = m.sigla;
    });
  }

  // Fichas de Datación (una por muestra DAT, hasta POOL_DAT)
  porTipo.DAT.forEach((m, k) => {
    const n = k + 1;
    if (n > POOL_DAT) return;
    rep[`FD${n}_solicitante`] = m.geologo;
    rep[`FD${n}_id`] = m.sigla;
    rep[`FD${n}_litologia`] = m.litologia;
    rep[`FD${n}_unidad`] = m.unidad;
    rep[`FD${n}_clasificacion`] = m.litologia;
  });

  return { rep, porTipo };
}

// Marca como ocultas (state="hidden") en workbook.xml las hojas que no se usan.
// openpyxl escribe state="visible" explícito, así que hay que reemplazarlo, no solo añadir.
function ocultarHojas(workbookXml, debeOcultar) {
  return workbookXml.replace(/<sheet\b[^>]*?\/>/g, (tag) => {
    const nm = (tag.match(/name="([^"]*)"/) || [])[1] || "";
    if (!debeOcultar(nm)) return tag;
    if (/state="[^"]*"/.test(tag)) return tag.replace(/state="[^"]*"/, 'state="hidden"');
    return tag.replace(/^<sheet\b/, '<sheet state="hidden"');
  });
}

async function generarMemo(seleccion) {
  const bytes = await cargarPlantilla();
  const zip = await JSZip.loadAsync(bytes);
  const { rep, porTipo } = construirReemplazos(seleccion);

  // Reemplazo de tokens en cada hoja (los @@...@@ sobrantes se blanquean).
  const sheetPaths = Object.keys(zip.files).filter((p) => /^xl\/worksheets\/sheet\d+\.xml$/.test(p));
  for (const p of sheetPaths) {
    let xml = await zip.file(p).async("string");
    xml = xml.replace(/@@([A-Za-z0-9_]+)@@/g, (m, tok) =>
      (tok in rep) ? escXml(rep[tok]) : "");
    zip.file(p, xml);
  }

  // Ocultar hojas Memo de tipos ausentes y fichas de datación sobrantes.
  const nDat = porTipo.DAT.length;
  const wbPath = "xl/workbook.xml";
  let wbXml = await zip.file(wbPath).async("string");
  wbXml = ocultarHojas(wbXml, (nm) => {
    const s = nm.trim();
    if (s.startsWith("Memo")) {
      if (s.includes("CT")) return porTipo.CT.length === 0;
      if (s.includes("U-Pb")) return porTipo.DAT.length === 0;
      if (s.includes("RX")) return porTipo.RX.length === 0;
    }
    if (s.startsWith("Ficha Datación")) {
      const idx = parseInt(s.replace(/[^0-9]/g, ""), 10);
      return !(idx >= 1 && idx <= nDat);
    }
    return false;
  });
  zip.file(wbPath, wbXml);

  const blob = await zip.generateAsync({
    type: "blob", compression: "DEFLATE",
    mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const d = new Date();
  const nombre = `Memo laboratorio ${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}.xlsx`;
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = nombre;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 4000);
  return nombre;
}

/* ================== UI ================== */
function App() {
  const [texto, setTexto] = useState("");
  const [muestras, setMuestras] = useState([]);
  const [aviso, setAviso] = useState(null);
  const [filtro, setFiltro] = useState("");
  const [ftipo, setFtipo] = useState("TODOS");
  const [generando, setGenerando] = useState(false);
  const fileRef = useRef(null);

  function cargar(txt) {
    const { rows, error } = parseTabla(txt);
    if (error) { setAviso({ tipo: "error", msg: error }); setMuestras([]); return; }
    setMuestras(rows);
    const sinTipo = rows.filter((r) => r.tipo === "?").length;
    setAviso({ tipo: "ok", msg: `Se cargaron ${rows.length} muestras.` + (sinTipo ? ` ${sinTipo} sin tipo detectado (revisar).` : "") });
  }
  function onDrop(e) {
    e.preventDefault();
    const f = e.dataTransfer.files[0]; if (!f) return;
    const rd = new FileReader();
    rd.onload = () => { setTexto(rd.result); cargar(rd.result); };
    rd.readAsText(f);
  }
  function onFile(e) {
    const f = e.target.files[0]; if (!f) return;
    const rd = new FileReader();
    rd.onload = () => { setTexto(rd.result); cargar(rd.result); };
    rd.readAsText(f);
  }

  const visibles = useMemo(() => muestras.map((m, i) => ({ ...m, _i: i }))
    .filter((m) => (ftipo === "TODOS" || m.tipo === ftipo))
    .filter((m) => m.sigla.toLowerCase().includes(filtro.toLowerCase())), [muestras, filtro, ftipo]);

  const seleccionadas = muestras.filter((m) => m.sel);
  const conteo = useMemo(() => {
    const c = { CT: 0, DAT: 0, RX: 0, "?": 0 };
    seleccionadas.forEach((m) => c[m.tipo]++);
    return c;
  }, [muestras]);

  function toggle(i) { setMuestras((ms) => ms.map((m, k) => k === i ? { ...m, sel: !m.sel } : m)); }
  function setUnidad(i, v) { setMuestras((ms) => ms.map((m, k) => k === i ? { ...m, unidad: v } : m)); }
  function setTipo(i, v) { setMuestras((ms) => ms.map((m, k) => k === i ? { ...m, tipo: v } : m)); }
  function todos(val) { setMuestras((ms) => ms.map((m) => visibles.some(v => v._i === ms.indexOf(m)) ? { ...m, sel: val } : m)); }

  function validar() {
    const errs = [];
    if (seleccionadas.length === 0) errs.push("No hay muestras seleccionadas.");
    if (seleccionadas.length > MAX_FI) errs.push(`La Ficha de Ingreso admite ${MAX_FI} muestras; hay ${seleccionadas.length} seleccionadas. Reduce la selección o genera en dos tandas.`);
    if (conteo.CT > MAX_LISTADO) errs.push(`El memo de cortes transparentes admite ${MAX_LISTADO} muestras; hay ${conteo.CT}.`);
    if (conteo.DAT > MAX_LISTADO) errs.push(`El memo de datación admite ${MAX_LISTADO} muestras; hay ${conteo.DAT}.`);
    if (conteo.RX > MAX_LISTADO) errs.push(`El memo de rayos X admite ${MAX_LISTADO} muestras; hay ${conteo.RX}.`);
    if (conteo.DAT > POOL_DAT) errs.push(`La plantilla trae ${POOL_DAT} fichas de datación; hay ${conteo.DAT} muestras de datación. Genera las de datación en dos tandas.`);
    if (conteo["?"] > 0) errs.push(`${conteo["?"]} muestra(s) seleccionada(s) sin tipo detectado: corrige el tipo en la tabla o desmárcalas.`);
    return errs;
  }

  async function onGenerar() {
    const errs = validar();
    if (errs.length) { setAviso({ tipo: "error", msg: errs.join(" ") }); return; }
    setGenerando(true); setAviso(null);
    try {
      const nombre = await generarMemo(seleccionadas);
      setAviso({ tipo: "ok", msg: `Memo generado: ${nombre}. Revísalo en Excel y agrega el número de memorando y la firma.` });
    } catch (e) {
      setAviso({ tipo: "error", msg: "Error al generar: " + e.message });
    } finally { setGenerando(false); }
  }

  return (
    <div className="wrap">
      <header>
        <h1>Generador de Memos de Laboratorio</h1>
        <p className="sub">Cuadrángulo Central Los Cipreses · SERNAGEOMIN — <span className="ver">__APP_VERSION__</span></p>
      </header>

      {muestras.length === 0 && (
        <section className="card">
          <h2>1 · Cargar muestras</h2>
          <p className="hint">En ArcGIS, selecciona las filas en la <b>tabla de atributos</b>, cópialas (Ctrl+C) y <b>pégalas</b> aquí abajo. También puedes arrastrar un archivo <b>CSV</b>.</p>
          <div className="drop" onDragOver={(e) => e.preventDefault()} onDrop={onDrop} onClick={() => fileRef.current.click()}>
            Arrastra un CSV aquí o haz clic para elegirlo
            <input ref={fileRef} type="file" accept=".csv,.txt,.tsv" onChange={onFile} hidden />
          </div>
          <textarea value={texto} onChange={(e) => setTexto(e.target.value)}
            placeholder="Pega aquí la tabla copiada desde ArcGIS (con encabezados)…" rows={7} />
          <button className="primary" onClick={() => cargar(texto)} disabled={!texto.trim()}>Cargar muestras</button>
        </section>
      )}

      {aviso && <div className={"aviso " + aviso.tipo}>{aviso.msg}</div>}

      {muestras.length > 0 && (
        <section className="card">
          <div className="row between">
            <h2>2 · Seleccionar muestras</h2>
            <button className="link" onClick={() => { setMuestras([]); setTexto(""); setAviso(null); }}>Cargar otras</button>
          </div>
          <div className="row toolbar">
            <input className="buscar" placeholder="Buscar sigla…" value={filtro} onChange={(e) => setFiltro(e.target.value)} />
            <select value={ftipo} onChange={(e) => setFtipo(e.target.value)}>
              <option value="TODOS">Todos los tipos</option>
              <option value="CT">Cortes transparentes</option>
              <option value="DAT">Datación</option>
              <option value="RX">Rayos X</option>
              <option value="?">Sin tipo</option>
            </select>
            <button className="link" onClick={() => todos(true)}>Marcar visibles</button>
            <button className="link" onClick={() => todos(false)}>Desmarcar visibles</button>
          </div>
          <div className="tablewrap">
            <table>
              <thead>
                <tr><th></th><th>Sigla</th><th>Tipo</th><th>Litología</th><th>Unidad geológica</th><th>Este</th><th>Norte</th></tr>
              </thead>
              <tbody>
                {visibles.map((m) => (
                  <tr key={m._i} className={m.sel ? "sel" : ""}>
                    <td><input type="checkbox" checked={m.sel} onChange={() => toggle(m._i)} /></td>
                    <td className="mono">{m.sigla}</td>
                    <td>
                      <select value={m.tipo} onChange={(e) => setTipo(m._i, e.target.value)} className={m.tipo === "?" ? "warn" : ""}>
                        <option value="CT">Cortes</option>
                        <option value="DAT">Datación</option>
                        <option value="RX">Rayos X</option>
                        <option value="?">?</option>
                      </select>
                    </td>
                    <td>{m.litologia}</td>
                    <td><input className="celda" value={m.unidad} onChange={(e) => setUnidad(m._i, e.target.value)} /></td>
                    <td className="mono">{m.este}</td>
                    <td className="mono">{m.norte}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div className="resumen">
            Seleccionadas: <b>{seleccionadas.length}</b> ·
            Cortes {conteo.CT} · Datación {conteo.DAT} · Rayos X {conteo.RX}
            {conteo["?"] > 0 && <span className="warn"> · Sin tipo {conteo["?"]}</span>}
          </div>
          <button className="primary big" onClick={onGenerar} disabled={generando}>
            {generando ? "Generando…" : "Generar Memo (.xlsx)"}
          </button>
          <p className="hint">Se genera un solo archivo con la Ficha de Ingreso, un memo por cada tipo presente, y una ficha de datación por muestra de datación. El número de memorando y la firma se completan a mano en Excel.</p>
        </section>
      )}
    </div>
  );
}

ReactDOM.createRoot(document.getElementById("root")).render(<App />);
