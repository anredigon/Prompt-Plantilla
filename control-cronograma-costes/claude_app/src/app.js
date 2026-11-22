/* App de control de obra (valor ganado) para claude.ai.
 * Guarda cada proyecto en la base de datos del artefacto (proyectos/<id>), lee adjuntos con IA (sample)
 * y reutiliza el motor (motor.js) y el informe (informe.js) de la app sin conexión.
 * Depende de: XLSX (SheetJS), pdfjsLib (pdf.js), calcularDatos/leerLibro (motor.js), render y utilidades (informe.js).
 */
{
const ST = { db: null, user: null, sample: null, downloads: null, escribir: false, proyectos: [], id: null, datos: null,
             vista: 'inicio', paso: 0, borrador: null, corteVista: null, tab: 'informe', semana: null };
const vista = document.getElementById('vista');
const aviso = document.getElementById('aviso');
const hoyISO = () => new Date().toISOString().slice(0, 10);
const nuevoId = () => 'p' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
const clon = o => JSON.parse(JSON.stringify(o));
const n = v => { if (typeof v === 'number') return isFinite(v) ? v : null; const t = String(v ?? '').trim().replace(/\s|€/g, ''); if (!t) return null;
  const s = t.includes(',') ? t.replace(/\./g, '').replace(',', '.') : t; const x = Number(s); return isFinite(x) ? x : null; };
const semanaDe = (inicioISO, fechaISO) => Math.floor((new Date(fechaISO + 'T00:00:00Z') - new Date(inicioISO + 'T00:00:00Z')) / 6048e5) + 1;
const lunesDe = iso => { const d = new Date(iso + 'T00:00:00Z'); const dia = (d.getUTCDay() + 6) % 7; return new Date(d - dia * 864e5).toISOString().slice(0, 10); };
function mensaje(txt, tipo) { aviso.textContent = txt; aviso.className = 'aviso ' + (tipo || ''); aviso.hidden = !txt; if (txt) setTimeout(() => { if (aviso.textContent === txt) aviso.hidden = true; }, 7000); }

// =================================================================== capacidades
async function iniciar() {
  const [db, user, sample, downloads] = await Promise.all(['db', 'user', 'sample', 'downloads'].map(c => window.claude ? claude.use(c).catch(() => null) : null));
  Object.assign(ST, { db, user, sample, downloads });
  ST.escribir = !!db && ((user && (await user.can('data.write').catch(() => null))) ?? true) !== false;
  document.getElementById('rol').textContent = !db ? 'Sin base de datos: ábrela en claude.ai' : ST.escribir ? 'Puedes editar' : 'Solo lectura';
  if (!db) { vista.innerHTML = '<div class="card"><h2>La base de datos no está disponible</h2><p class="muted">Abre la app desde claude.ai con tu cuenta para ver y guardar proyectos.</p></div>'; return; }
  const espera = setTimeout(() => mensaje('La base de datos tarda en responder; los proyectos aparecerán al conectar.', 'warn'), 10000);
  db.collection('proyectos').onSnapshot(snap => {
    clearTimeout(espera);
    ST.proyectos = snap.docs.map(d => ({ id: d.id, ...d.data() }));
    // Cambios de otros usuarios: se reflejan salvo mientras se rellena un formulario (no se pierde lo escrito)
    const ocupado = ST.vista === 'asistente' || (ST.vista === 'proyecto' && ST.tab === 'semana');
    if (ST.id && !ocupado) { const p = ST.proyectos.find(x => x.id === ST.id); if (p) ST.datos = clon(p.datos); }
    if (!ocupado) pintar();
  }, e => { clearTimeout(espera); mensaje('No se pueden leer los proyectos (' + e.code + ').', 'error'); });
  pintar();
}

async function guardar(id, datos) {
  let res = null;
  try { res = calcularDatos(datos, {}); } catch (e) { /* proyecto incompleto: se guarda igualmente */ }
  const resumen = res ? { BAC: res.D.BAC, EAC: res.D.EAC, SPI: res.D.SPI, CPI: res.D.CPI, fin_prev: res.D.fin_prev, alertas: res.D.alertas.filter(a => a.nivel === 'critical').length } : null;
  try {
    await ST.db.doc('proyectos/' + id).set({ nombre: datos.proyecto, actualizado: new Date().toISOString(), corte: datos.corte, n_sem: datos.n_sem, resumen, datos });
    mensaje('Guardado.', 'ok'); return true;
  } catch (e) {
    mensaje(e.code === 'invalid_argument' ? 'No tienes permiso para editar o el proyecto es demasiado grande para guardarse.' : 'No se ha podido guardar (' + e.code + '). Inténtalo de nuevo.', 'error');
    if (e.code === 'invalid_argument') ST.escribir = false;
    return false;
  }
}

// =================================================================== lectura de adjuntos
async function leerArchivos(files) {
  const partes = [], imagenes = [];
  for (const f of files) {
    const nombre = f.name.toLowerCase();
    try {
      if (nombre.endsWith('.pdf')) {
        const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(await f.arrayBuffer()) }).promise;
        let texto = '';
        for (let i = 1; i <= Math.min(pdf.numPages, 40); i++) {
          const c = await (await pdf.getPage(i)).getTextContent();
          const lineas = {};
          for (const it of c.items) { const y = Math.round(it.transform[5]); (lineas[y] = lineas[y] || []).push([it.transform[4], it.str]); }
          texto += `\n--- Página ${i} ---\n` + Object.keys(lineas).sort((a, b) => b - a).map(y => lineas[y].sort((a, b) => a[0] - b[0]).map(x => x[1]).join('  ')).join('\n');
        }
        if (texto.replace(/\s|---|Página \d+/g, '').length < 200) {  // PDF escaneado: se envían las páginas como imagen
          for (let i = 1; i <= Math.min(pdf.numPages, 4); i++) {
            const pg = await pdf.getPage(i), vp = pg.getViewport({ scale: 2 });
            const cv = Object.assign(document.createElement('canvas'), { width: vp.width, height: vp.height });
            await pg.render({ canvasContext: cv.getContext('2d'), viewport: vp }).promise;
            imagenes.push(await new Promise(r => cv.toBlob(r, 'image/png')));
          }
          partes.push(`[${f.name}: PDF escaneado, ${Math.min(pdf.numPages, 4)} página(s) adjuntas como imagen]`);
        } else partes.push(`[Archivo ${f.name}]${texto}`);
      } else if (/\.(xlsx|xlsm|xls|csv|txt|tsv)$/.test(nombre)) {
        const wb = XLSX.read(new Uint8Array(await f.arrayBuffer()), { type: 'array' });
        for (const h of wb.SheetNames) {
          const csv = XLSX.utils.sheet_to_csv(wb.Sheets[h], { blankrows: false });
          if (csv.trim()) partes.push(`[Archivo ${f.name} · hoja ${h}]\n${csv}`);
        }
      } else if (/^image\//.test(f.type)) { imagenes.push(f); partes.push(`[${f.name}: imagen adjunta]`); }
      else partes.push(`[${f.name}: formato no soportado, se ignora]`);
    } catch (e) { partes.push(`[${f.name}: no se ha podido leer (${e.message || e})]`); }
  }
  let texto = partes.join('\n\n');
  if (texto.length > 150000) texto = texto.slice(0, 150000) + '\n[... recortado ...]';
  return { texto, imagenes };
}

let ctlIA = null;
async function pedirIA(boton, instruccion, files, textoUsuario) {
  if (!ST.sample) { mensaje('La IA no está disponible en esta vista.', 'error'); return null; }
  const estado = boton.parentElement.querySelector('.estado-ia');
  boton.disabled = true; estado.innerHTML = 'Leyendo adjuntos… <button type="button" class="link" data-accion="parar-ia">Parar</button>';
  try {
    const { texto, imagenes } = await leerArchivos(files || []);
    if (!texto.trim() && !textoUsuario.trim() && !imagenes.length) { mensaje('Adjunta un archivo o escribe los datos primero.', 'warn'); return null; }
    if (imagenes.length) {
      const lim = await ST.sample.limits().catch(() => null);
      if (!lim || !lim.images) { mensaje('Esta vista no puede enviar imágenes a la IA: adjunta un PDF con texto, un Excel o escribe los datos.', 'warn'); imagenes.length = 0; }
      else imagenes.splice(lim.images.maxCount);
    }
    estado.innerHTML = 'La IA está leyendo (puede tardar 30–90 s)… <button type="button" class="link" data-accion="parar-ia">Parar</button>';
    ctlIA = new AbortController();
    const prompt = instruccion + '\n\n=== DATOS ESCRITOS POR EL USUARIO ===\n' + (textoUsuario || '(ninguno)') + '\n\n=== CONTENIDO DE LOS ADJUNTOS ===\n' + (texto || '(ninguno)');
    const r = await ST.sample.json(prompt, { signal: ctlIA.signal, images: imagenes.length ? imagenes : undefined, cache: false });
    estado.textContent = 'Hecho. Revisa y corrige la tabla antes de continuar.';
    return r;
  } catch (e) {
    const txt = { cancelled: 'Lectura detenida.', not_granted: 'No has permitido que la app use Claude.', rate_limited: 'Demasiadas consultas o límite de uso alcanzado: espera un poco.',
                  invalid_json: 'La IA no devolvió datos legibles. Prueba de nuevo o con menos páginas.', prompt_too_large: 'El adjunto es demasiado grande: envía solo las páginas relevantes.',
                  refused: 'La IA no ha procesado este contenido.' }[e.code] || ('No se ha podido leer (' + (e.code || e.message || e) + ').');
    estado.textContent = txt; return null;
  } finally { boton.disabled = false; ctlIA = null; }
}

const PROMPTS = {
  presupuesto: `Eres técnico de mediciones y presupuestos de obra. Extrae del documento TODAS las partidas con importe.
Responde SOLO con JSON: {"partidas":[{"cod":"P01","capitulo":"","descripcion":"","cantidad":0,"unidad":"ud","importe":0}],"extras":[{"descripcion":"","importe":0}],"total_documento":0,"notas":""}
Reglas: "importe" es el TOTAL de la línea en euros como número (punto decimal, sin separador de miles). No inventes partidas ni importes.
Si el documento no trae código, numera P01, P02... por orden. "capitulo" es el encabezado de la sección a la que pertenece.
Las líneas marcadas como "no incluido", "extra", "opcional" o fuera del total van en "extras", no en "partidas".
"total_documento" es el total que declara el documento (o null). En "notas" avisa de lo dudoso (subtotales, líneas ilegibles).`,
  cronograma: inicio => `Eres planificador de obra. Extrae las actividades del cronograma adjunto (puede venir de MS Project, Primavera o Excel).
La semana 1 empieza el lunes ${inicio}. Convierte fechas a número de semana: semana = floor((fecha - ${inicio}) / 7 días) + 1.
Responde SOLO con JSON: {"actividades":[{"id":"ACT-010","nombre":"","sem_inicio":1,"sem_fin":4,"predecesora":""}],"notas":""}
Reglas: incluye solo actividades de ejecución (no hitos de duración cero ni resúmenes si hay detalle). Si no hay ID, numera ACT-010, ACT-020...
"predecesora" es el id de una predecesora principal o "". No inventes fechas; si faltan, dilo en "notas".`,
  proponerCronograma: (partidas, nSem) => `Eres planificador de obra. Propón un cronograma realista de ${nSem} semanas para ejecutar este presupuesto, agrupando partidas en actividades (entre 6 y 20) con una secuencia lógica de obra (suministros antes de montaje, pruebas y puesta en marcha al final).
Presupuesto (código | capítulo | descripción | importe):
${partidas.map(p => `${p.cod} | ${p.cap} | ${p.desc} | ${p.bac}`).join('\n')}
Responde SOLO con JSON: {"actividades":[{"id":"ACT-010","nombre":"","sem_inicio":1,"sem_fin":4,"predecesora":""}],"notas":"supuestos de la propuesta"}`,
  medicion: (partidas, actividades) => `Eres jefe de control de obra. Para cada partida decide cómo se mide su avance y a qué actividad pertenece.
Fuentes de avance posibles:
- ALBARAN: suministro de equipos o materiales que se gana al entregarse (compresores, cuadros, evaporadores...). Perfil "ENTREGA".
- REVIT: elementos modelados que se ganan al instalarse en el modelo (tuberías en m, soportes, desagües, unidades montadas). Perfil "LINEAL".
- MANUAL: lo que no está en el modelo (cargas, pruebas, protecciones, legalización, puesta en marcha). Perfil "LINEAL".
- DERIVADO: mano de obra o montaje global, que avanza con las horas de los elementos instalados en el modelo. Perfil "LINEAL".
Para REVIT y ALBARAN da la unidad de medición y la cantidad: usa la de la partida; si la partida es "1 ud global" pero se mide en el modelo (p. ej. tuberías o aislamiento en metros), pon la unidad real ("m") y cantidad null para que la escriba el usuario.
Para cada partida que se monte en obra (equipos y elementos del modelo) estima horas-hombre de montaje por unidad (o por metro) en "horas"; para el resto null.
Actividades: ${actividades.map(a => `${a.id} ${a.nombre} (S${a.si}-S${a.sf})`).join('; ')}
Partidas (código | capítulo | descripción | cantidad | unidad | importe):
${partidas.map(p => `${p.cod} | ${p.cap} | ${p.desc} | ${p.cant ?? ''} | ${p.ud} | ${p.bac}`).join('\n')}
Responde SOLO con JSON: {"asignaciones":[{"cod":"P01","act":"ACT-010","fuente":"ALBARAN","perfil":"ENTREGA","ud":"ud","cmed":2,"horas":120}],"dudas":["..."]}`,
  albaranes: (partidas, corte) => `Eres administrativo de obra. Extrae del/los albarán(es) de entrega las cantidades entregadas y asígnalas a las partidas del presupuesto.
Partidas (código | descripción | unidad): ${partidas.map(p => `${p.cod} | ${p.desc} | ${p.ud}`).join(' ; ')}
Responde SOLO con JSON: {"lineas":[{"cod":"P01","cantidad":1,"albaran":"nº","fecha":"AAAA-MM-DD o null","texto_original":"","confianza":"alta|media|baja"}],"notas":""}
Si una línea no corresponde a ninguna partida, pon "cod":"" y explícalo en "notas". No inventes cantidades. Semana actual de obra: ${corte}.`,
  planilla: partidas => `Eres técnico de certificaciones. Extrae las líneas de la certificación/planilla del contratista y asígnalas a las partidas del presupuesto.
Partidas (código | descripción | unidad | importe presupuesto): ${partidas.map(p => `${p.cod} | ${p.desc} | ${p.ud} | ${p.bac}`).join(' ; ')}
Responde SOLO con JSON: {"numero":"","fecha":null,"lineas":[{"cod":"P01","cantidad":0,"precio":0,"importe":0,"texto_original":"","observacion":""}],"notas":""}
Reglas: "cantidad" es la cantidad certificada EN ESTE periodo (no la acumulada a origen si viene desglosada; si solo hay acumulado, dilo en "notas").
"precio" es el precio unitario aplicado; si solo hay importe, precio = importe / cantidad. Si un precio difiere del presupuesto, anótalo en "observacion". No inventes.`,
};

// =================================================================== Revit (archivo del botón pyRevit o tabla exportada)
async function leerRevit(file) {
  const wb = XLSX.read(new Uint8Array(await file.arrayBuffer()), { type: 'array' });
  const filas = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: true, defval: null });
  const norm = s => String(s || '').toUpperCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^A-Z0-9]/g, '');
  const cab = (filas[0] || []).map(norm);
  const col = (nombres, porDefecto) => { for (const x of nombres) { const i = cab.indexOf(norm(x)); if (i >= 0) return i; } return porDefecto; };
  const c = { cod: col(['CC_COD_PARTIDA'], 4), q: col(['CANTIDAD', 'Cantidad', 'Length', 'Longitud'], 6), estado: col(['CC_ESTADO'], 8), sem: col(['CC_SEMANA_INSTALADO'], 9) };
  const agreg = {};
  let total = 0;
  for (const r of filas.slice(1)) {
    const cod = String(r[c.cod] ?? '').trim();
    if (!cod) continue;
    total++;
    const est = String(r[c.estado] ?? '').trim().toLowerCase() === 'instalado' ? 'Instalado' : 'Pendiente';
    const sem = n(r[c.sem]);
    const k = cod + '|' + est + '|' + (est === 'Instalado' && sem != null ? Math.round(sem) : '');
    agreg[k] = agreg[k] || { cod, q: 0, estado: est, sem: est === 'Instalado' && sem != null ? Math.round(sem) : null, nElem: 0 };
    agreg[k].q += n(r[c.q]) ?? 1; agreg[k].nElem++;
  }
  // Agregado por partida/estado/semana: el motor da el mismo resultado y el proyecto cabe en la base de datos
  return { items: Object.values(agreg).map(x => ({ ...x, q: Math.round(x.q * 1000) / 1000 })), total, archivo: file.name, importado: new Date().toISOString() };
}

// =================================================================== exportaciones
async function guardarArchivo(nombre, data) {
  if (!ST.downloads) { mensaje('Las descargas no están disponibles en esta vista.', 'error'); return; }
  try { await ST.downloads.save({ filename: nombre, data }); }
  catch (e) { if (e.code !== 'declined') mensaje('No se ha podido descargar (' + e.code + ').', 'error'); }
}
const slug = s => (s || 'proyecto').normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^\w-]+/g, '_').replace(/^_+|_+$/g, '').slice(0, 60);
const xlsxBlob = wb => new Blob([XLSX.write(wb, { bookType: 'xlsx', type: 'array' })]);

function excelProyecto(d) {  // misma estructura que la plantilla: se puede abrir con la app sin conexión
  const wb = XLSX.utils.book_new();
  const hoja = (nombre, filas) => XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(filas), nombre);
  const inicioSerial = (new Date(d.inicio + 'T00:00:00Z') - Date.UTC(1899, 11, 30)) / 864e5;
  hoja('Parametros', [['Parámetros del proyecto'], [], ['Proyecto', d.proyecto], ['Fecha inicio (lunes semana 1)', inicioSerial], ['Nº semanas', d.n_sem],
       ['Semana de corte', d.corte], ['Fecha de corte', inicioSerial + 7 * d.corte - 1], ['Contingencia', d.cont_pct], ['Umbral total', d.umbral], ['Umbral actividad', d.umbral_act]]);
  wb.Sheets.Parametros.B4.z = wb.Sheets.Parametros.B7.z = 'dd/mm/yyyy';
  hoja('Cronograma', [['ID_ACTIVIDAD', 'ACTIVIDAD', 'SEM_INICIO', 'SEM_FIN', '', '', '', 'PREDECESORA'], ...d.actividades.map(a => [a.id, a.nombre, a.si, a.sf, '', '', '', a.pred || ''])]);
  hoja('Oferta', [['COD_PARTIDA', 'CAPÍTULO', 'DESCRIPCIÓN', 'CANT. OFERTA', 'IMPORTE OFERTA €', 'BAC €', 'ID_ACTIVIDAD', 'FUENTE AVANCE', 'UD MEDICIÓN', 'CANT. MEDICIÓN',
       ...Array(12).fill(''), 'PERFIL PV', 'SEM. ENTREGA PREVISTA'],
       ...d.partidas.map(p => [p.cod, p.cap, p.desc, p.cant ?? '', p.bac, p.bac, p.act, p.fuente, p.ud, p.cmed ?? '', ...Array(12).fill(''), p.perfil, p.sem_ent ?? ''])]);
  hoja('Revit_Export', [['ElementId', 'Categoría', 'Familia : Tipo', 'Zona / Nivel', 'CC_COD_PARTIDA', 'CC_ID_ACTIVIDAD', 'CANTIDAD', 'UD', 'CC_ESTADO', 'CC_SEMANA_INSTALADO'],
       ...(d.revit || []).map(e => ['', '', `Agregado (${e.nElem || 1} elementos)`, '', e.cod, '', e.q, '', e.estado, e.sem ?? ''])]);
  hoja('Suministros', [['SEMANA', 'COD_PARTIDA', 'CANTIDAD ENTREGADA', 'Nº ALBARÁN', 'OBSERVACIÓN'], ...(d.suministros || []).map(x => [x.sem, x.cod, x.q, x.alb || '', ''])]);
  hoja('Horas_Montaje', [['COD_PARTIDA', 'DESCRIPCIÓN', 'HORAS POR UD', 'UD'], ...Object.entries(d.horas || {}).map(([c, h]) => [c, (d.partidas.find(p => p.cod === c) || {}).desc || '', h, ''])]);
  hoja('Avance_Manual', [['COD_PARTIDA', 'DESCRIPCIÓN', ...Array.from({ length: 104 }, (_, i) => 'S' + (i + 1))],
       ...Object.entries(d.manual || {}).map(([c, s]) => [c, (d.partidas.find(p => p.cod === c) || {}).desc || '', ...s.map(v => v ?? '')])]);
  hoja('Planillas', [['Nº PLANILLA', 'SEMANA', 'COD_PARTIDA', 'CANT. CERTIFICADA', 'PRECIO APLICADO €', 'IMPORTE €', 'OBSERVACIÓN', 'SEMANA PAGO'],
       ...(d.planillas || []).map(x => [x.n, x.sem, x.cod, x.q, x.pu, Math.round(x.q * x.pu * 100) / 100, x.obs || '', x.sem_pago ?? ''])]);
  hoja('Extras', [['COD', 'DESCRIPCIÓN', 'IMPORTE €', 'ESTADO', 'SEMANA LÍMITE DECISIÓN', 'MOTIVO DEL LÍMITE'], ...(d.extras || []).map(e => [e.cod, e.desc, e.imp, e.estado, e.limite ?? '', e.motivo || ''])]);
  hoja('Contingencia', [['SEMANA', 'CONCEPTO', 'IMPORTE €', 'ESTADO'], ...(d.contingencia || []).map(c => [c.sem, c.concepto, c.imp, c.estado])]);
  return wb;
}

async function informeHTML(D) {
  const [css, informe] = await Promise.all([document.getElementById('estilos-informe').textContent, fetch('./informe.js').then(r => r.text())]);
  const json = JSON.stringify(D).replace(/</g, '\\u003c');
  return '<!DOCTYPE html>\n<html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>Informe · ' + esc(D.proyecto) + '</title><style>' + css +
    '</style></head><body><div class="viz-root"><div class="wrap" id="app"></div><div class="tt" id="tt" role="tooltip"></div></div><script>' + informe.replace(/<\/script/gi, '<\\/script') +
    '\nrender(' + json + ', "informe");\ndocument.addEventListener("click", e => { if (e.target.id === "theme") { const r = document.documentElement; r.dataset.theme = r.dataset.theme === "dark" ? "light" : "dark"; } });<\/script></body></html>';
}

// =================================================================== pintado general
function pintar() {
  document.body.dataset.vista = ST.vista;
  if (ST.vista === 'inicio') return pintarInicio();
  if (ST.vista === 'asistente') return pintarAsistente();
  if (ST.vista === 'proyecto') return pintarProyecto();
}

function pintarInicio() {
  const lista = [...ST.proyectos].sort((a, b) => String(b.actualizado).localeCompare(String(a.actualizado)));
  vista.innerHTML = `
  <div class="cabecera-vista"><div><h1>Proyectos</h1><p class="muted">Control semanal de cronograma y costes con valor ganado.</p></div>
    ${ST.escribir ? `<div class="botones"><button class="primario" data-accion="nuevo">Nuevo proyecto</button>
      <label class="boton" for="importar">Importar Excel de la plantilla</label></div>` : ''}</div>
  ${lista.length ? `<div class="rejilla">${lista.map(p => {
    const r = p.resumen; const nivel = r && r.SPI != null ? nivelPor(Math.min(r.SPI, r.CPI ?? 9), 0.95) : 'sin_iniciar';
    return `<button class="card tarjeta" data-accion="abrir" data-id="${esc(p.id)}">
      <b>${esc(p.nombre)}</b><span class="muted">Semana ${p.corte} de ${p.n_sem} · actualizado ${fecha(String(p.actualizado).slice(0, 10))}</span>
      ${r ? `<span>EAC ${eur(r.EAC)} <span class="muted">(oferta ${eur(r.BAC)})</span></span><span>${badge(nivel)} ${r.alertas ? `<span class="muted">· ${r.alertas} alerta(s) crítica(s)</span>` : ''}</span>` : '<span class="muted">Configuración incompleta</span>'}
    </button>`; }).join('')}</div>`
  : `<div class="card vacio"><h2>Aún no hay proyectos</h2><p>${ST.escribir ? 'Pulsa <b>Nuevo proyecto</b>: te pediré los datos paso a paso y podrás adjuntar la oferta, el cronograma y lo que tengas.' : 'Cuando el responsable cree un proyecto aparecerá aquí.'}</p></div>`}`;
}

// =================================================================== asistente de proyecto nuevo / configuración
const PASOS = ['Datos generales', 'Presupuesto', 'Cronograma', 'Cómo se mide', 'Extras y contingencia', 'Revisar y guardar'];
function nuevoBorrador() {
  const lunes = lunesDe(hoyISO());
  return { proyecto: '', inicio: lunes, n_sem: 24, corte: 1, cont_pct: 0, umbral: 0.95, umbral_act: 0.9, actividades: [], partidas: [], revit: [], manual: {},
           suministros: [], horas: {}, planillas: [], extras: [], contingencia: [], avisos: [], total_control: null };
}

function tabla(id, columnas, filas, opciones) {
  opciones = opciones || {};
  return `<div class="scroll"><table class="editable" id="${id}"><thead><tr>${columnas.map(c => `<th class="${c.num ? 'n' : ''}">${c.t}</th>`).join('')}${opciones.borrar === false ? '' : '<th></th>'}</tr></thead><tbody>
    ${filas.map((f, i) => `<tr data-i="${i}">${columnas.map(c => {
      const v = f[c.k] ?? '';
      if (c.opciones) return `<td><select data-k="${c.k}">${c.opciones.map(o => `<option ${String(o) === String(v) ? 'selected' : ''}>${esc(o)}</option>`).join('')}</select></td>`;
      return `<td class="${c.num ? 'n' : ''}"><input data-k="${c.k}" value="${esc(v)}" ${c.num ? 'inputmode="decimal"' : ''} ${c.ancho ? `style="width:${c.ancho}"` : ''}></td>`; }).join('')}
      ${opciones.borrar === false ? '' : `<td><button class="link" data-accion="borrar-fila" data-tabla="${id}" data-i="${i}" title="Quitar fila">✕</button></td>`}</tr>`).join('')}
  </tbody></table></div>${opciones.anadir === false ? '' : `<button class="link" data-accion="anadir-fila" data-tabla="${id}">+ Añadir fila</button>`}`;
}
const zonaIA = (id, texto, botones) => `<div class="zona-ia">
  <label class="drop" data-drop="${id}"><input type="file" multiple id="adj-${id}" hidden accept=".pdf,.xlsx,.xls,.xlsm,.csv,.txt,image/*">
    <b>Adjunta archivos</b><span class="muted">PDF, Excel, CSV o fotos. Arrastra aquí o pulsa.</span><span class="lista-adj" id="lista-${id}"></span></label>
  <textarea id="txt-${id}" rows="3" placeholder="${esc(texto)}"></textarea>
  <div class="acciones-ia">${botones}<span class="estado-ia"></span></div></div>`;

// columnas de cada tabla editable y el lugar del borrador donde viven
const TABLAS = {
  't-partidas': { lista: d => d.partidas, cols: [{ k: 'cod', t: 'Código', ancho: '70px' }, { k: 'cap', t: 'Capítulo' }, { k: 'desc', t: 'Descripción', ancho: '260px' }, { k: 'cant', t: 'Cant.', num: 1, ancho: '70px' }, { k: 'ud', t: 'Ud', ancho: '50px' }, { k: 'bac', t: 'Importe €', num: 1, ancho: '110px' }], nueva: () => ({ cod: '', cap: '', desc: '', cant: null, ud: 'ud', bac: null }) },
  't-extras': { lista: d => d.extras, cols: [{ k: 'cod', t: 'Código', ancho: '60px' }, { k: 'desc', t: 'Descripción', ancho: '260px' }, { k: 'imp', t: 'Importe €', num: 1, ancho: '110px' }, { k: 'estado', t: 'Estado', opciones: ['Sin decidir', 'Aprobado', 'Rechazado'] }, { k: 'limite', t: 'Semana límite', num: 1, ancho: '70px' }, { k: 'motivo', t: 'Motivo del límite', ancho: '220px' }], nueva: d => ({ cod: 'X' + String(d.extras.length + 1).padStart(2, '0'), desc: '', imp: null, estado: 'Sin decidir', limite: null, motivo: '' }) },
  't-actividades': { lista: d => d.actividades, cols: [{ k: 'id', t: 'ID', ancho: '80px' }, { k: 'nombre', t: 'Actividad', ancho: '260px' }, { k: 'si', t: 'Semana inicio', num: 1, ancho: '70px' }, { k: 'sf', t: 'Semana fin', num: 1, ancho: '70px' }, { k: 'pred', t: 'Predecesora', ancho: '80px' }], nueva: d => ({ id: 'ACT-' + String((d.actividades.length + 1) * 10).padStart(3, '0'), nombre: '', si: 1, sf: 1, pred: '' }) },
  't-medicion': { lista: d => d.partidas, cols: null, nueva: null },
  't-contingencia': { lista: d => d.contingencia, cols: [{ k: 'sem', t: 'Semana', num: 1, ancho: '60px' }, { k: 'concepto', t: 'Concepto', ancho: '300px' }, { k: 'imp', t: 'Importe €', num: 1, ancho: '110px' }, { k: 'estado', t: 'Estado', opciones: ['Propuesto', 'Aprobado', 'Rechazado'] }], nueva: d => ({ sem: d.corte, concepto: '', imp: null, estado: 'Propuesto' }) },
  't-alb': { lista: () => ST.semana.albaranes, cols: null, nueva: () => ({ cod: '', q: null, alb: '', sem: ST.semana.corte }) },
  't-pla': { lista: () => ST.semana.planillas, cols: null, nueva: () => ({ n: '', sem: ST.semana.corte, cod: '', q: null, pu: null, obs: '' }) },
};
const colsMedicion = d => [{ k: 'cod', t: 'Partida', ancho: '60px' }, { k: 'desc', t: 'Descripción', ancho: '200px' },
  { k: 'act', t: 'Actividad', opciones: ['', ...d.actividades.map(a => a.id)] }, { k: 'fuente', t: 'Se mide por', opciones: ['ALBARAN', 'REVIT', 'MANUAL', 'DERIVADO'] },
  { k: 'perfil', t: 'Planificado', opciones: ['ENTREGA', 'LINEAL'] }, { k: 'ud', t: 'Ud medición', ancho: '55px' }, { k: 'cmed', t: 'Cant. medición', num: 1, ancho: '80px' },
  { k: 'sem_ent', t: 'Semana entrega', num: 1, ancho: '60px' }, { k: 'horas', t: 'Horas montaje / ud', num: 1, ancho: '70px' }];
const colsAlb = () => [{ k: 'sem', t: 'Semana', num: 1, ancho: '60px' }, { k: 'cod', t: 'Partida', opciones: ['', ...ST.datos.partidas.map(p => p.cod)] }, { k: 'q', t: 'Cantidad', num: 1, ancho: '80px' }, { k: 'alb', t: 'Nº albarán', ancho: '110px' }, { k: 'texto', t: 'Texto original', ancho: '220px' }];
const colsPla = () => [{ k: 'n', t: 'Nº planilla', ancho: '80px' }, { k: 'sem', t: 'Semana', num: 1, ancho: '60px' }, { k: 'cod', t: 'Partida', opciones: ['', ...ST.datos.partidas.map(p => p.cod)] }, { k: 'q', t: 'Cantidad', num: 1, ancho: '80px' }, { k: 'pu', t: 'Precio €', num: 1, ancho: '100px' }, { k: 'obs', t: 'Observación', ancho: '220px' }];

function leerTablas() {  // vuelca lo escrito en las tablas visibles al borrador / semana
  for (const t of document.querySelectorAll('table.editable')) {
    const def = TABLAS[t.id]; if (!def) continue;
    const d = ST.vista === 'asistente' ? ST.borrador : ST.datos;
    const lista = def.lista(d);
    for (const tr of t.tBodies[0].rows) {
      const fila = lista[+tr.dataset.i]; if (!fila) continue;
      for (const el of tr.querySelectorAll('[data-k]')) {
        const k = el.dataset.k, numerica = el.closest('td').classList.contains('n');
        if (t.id === 't-medicion' && k === 'horas') { const h = n(el.value); if (h == null) delete d.horas[fila.cod]; else d.horas[fila.cod] = h; continue; }
        fila[k] = numerica ? n(el.value) : el.value.trim();
      }
    }
  }
}

function pintarAsistente() {
  const d = ST.borrador, p = ST.paso;
  const total = d.partidas.reduce((s, x) => s + (n(x.bac) || 0), 0);
  let cuerpo = '';
  if (p === 0) cuerpo = `
    <div class="form">
      <label>Nombre del proyecto<input id="f-nombre" value="${esc(d.proyecto)}" placeholder="p. ej. Instalación frigorífica nave 3"></label>
      <label>Inicio de obra<input type="date" id="f-inicio" value="${esc(d.inicio)}"><span class="muted">Se toma el lunes de esa semana como semana 1.</span></label>
      <label>Duración planificada (semanas)<input id="f-nsem" inputmode="numeric" value="${d.n_sem}"></label>
      <label>Contingencia aprobada (% sobre la oferta)<input id="f-cont" inputmode="decimal" value="${(d.cont_pct * 100).toLocaleString('es-ES')}"><span class="muted">0 si dirección no ha confirmado una contingencia.</span></label>
      <details><summary>Umbrales de alerta</summary>
        <label>Proyecto (SPI/CPI)<input id="f-umbral" inputmode="decimal" value="${String(d.umbral).replace('.', ',')}"></label>
        <label>Por actividad<input id="f-umbral-act" inputmode="decimal" value="${String(d.umbral_act).replace('.', ',')}"></label></details>
    </div>`;
  if (p === 1) cuerpo = `
    <p>Adjunta la <b>oferta o presupuesto</b> (PDF, Excel o una foto) y la IA extraerá las partidas. También puedes escribirlas o pegarlas.</p>
    ${zonaIA('pres', 'Opcional: pega aquí el presupuesto o añade indicaciones (p. ej. "el total firmado es 1.375.000 €").', '<button class="primario" data-accion="ia-presupuesto">Leer con IA</button>')}
    <h3>Partidas <span class="muted">(${d.partidas.length} · total ${eur(total)})</span></h3>
    ${tabla('t-partidas', TABLAS['t-partidas'].cols, d.partidas)}
    <label class="control">Importe total firmado (control)<input id="f-total" inputmode="decimal" value="${d.total_control ?? ''}">
      ${d.total_control != null ? (Math.abs(d.total_control - total) < 0.01 ? badge('good', 'Cuadra') : badge('critical', 'No cuadra: diferencia ' + eur(total - d.total_control))) : ''}</label>
    <h3>No incluido en la oferta (extras)</h3>${tabla('t-extras', TABLAS['t-extras'].cols, d.extras)}`;
  if (p === 2) cuerpo = `
    <p>Adjunta el <b>cronograma</b> (exportación de MS Project o Primavera en Excel/PDF) o pide a la IA una propuesta a partir del presupuesto.</p>
    ${zonaIA('crono', 'Opcional: describe la planificación (p. ej. "suministro de equipos semanas 3-8, montaje 4-18, pruebas 20-24").', '<button class="primario" data-accion="ia-cronograma">Leer con IA</button><button data-accion="ia-proponer-crono">Proponer desde el presupuesto</button>')}
    <h3>Actividades (semanas desde el ${fecha(d.inicio)})</h3>${tabla('t-actividades', TABLAS['t-actividades'].cols, d.actividades)}`;
  if (p === 3) cuerpo = `
    <p>Para cada partida: a qué actividad pertenece y <b>cómo se mide su avance</b>. Pide una propuesta a la IA y corrígela.</p>
    <ul class="leyenda-fuentes"><li><b>ALBARAN</b>: equipos y materiales; se gana al entregarse.</li><li><b>REVIT</b>: elementos del modelo; se gana al marcarlos como instalados en Revit.</li>
      <li><b>MANUAL</b>: lo que no está en el modelo; % semanal.</li><li><b>DERIVADO</b>: montaje global; avanza con las horas de lo instalado.</li></ul>
    <div class="acciones-ia"><button class="primario" data-accion="ia-medicion">Proponer con IA</button><span class="estado-ia"></span></div>
    ${tabla('t-medicion', colsMedicion(d), d.partidas.map(x => ({ ...x, horas: d.horas[x.cod] ?? '' })), { borrar: false, anadir: false })}
    <p class="muted">Las horas de montaje son una estimación: pide al instalador las reales. "Semana entrega" vacía = inicio de la actividad.</p>`;
  if (p === 4) cuerpo = `
    <h3>Extras fuera de oferta</h3>${tabla('t-extras', TABLAS['t-extras'].cols, d.extras)}
    <h3>Movimientos de contingencia</h3>${tabla('t-contingencia', TABLAS['t-contingencia'].cols, d.contingencia)}`;
  if (p === 5) {
    let res = null, error = '';
    try { res = calcularDatos(d, {}); } catch (e) { error = e.message; }
    cuerpo = error ? `<div class="error">${esc(error)}</div><p>Vuelve al paso correspondiente para completarlo.</p>`
      : `<div class="resumen-final"><div><b>${esc(d.proyecto || 'Sin nombre')}</b><span class="muted">${d.partidas.length} partidas · ${d.actividades.length} actividades · ${d.n_sem} semanas desde el ${fecha(d.inicio)}</span></div>
         <div>Oferta <b>${eur(res.D.BAC)}</b> · extras sin decidir <b>${eur(res.D.ext_pend)}</b> · contingencia <b>${pct(d.cont_pct)}</b></div></div>
         ${res.avisos.length ? `<div class="avisos"><b>Revisa:</b><ul>${res.avisos.map(a => `<li>${esc(a)}</li>`).join('')}</ul></div>` : badge('good', 'Sin avisos')}`;
  }
  vista.innerHTML = `
  <div class="cabecera-vista"><div><p class="muted">${ST.id ? 'Configuración del proyecto' : 'Nuevo proyecto'} · paso ${p + 1} de ${PASOS.length}</p><h1>${PASOS[p]}</h1></div>
    <button data-accion="cancelar">Cancelar</button></div>
  <ol class="pasos-barra">${PASOS.map((t, i) => `<li class="${i === p ? 'actual' : i < p ? 'hecho' : ''}">${t}</li>`).join('')}</ol>
  <div class="card cuerpo-paso">${cuerpo}</div>
  <div class="navegacion">${p > 0 ? '<button data-accion="anterior">Anterior</button>' : '<span></span>'}
    ${p < PASOS.length - 1 ? '<button class="primario" data-accion="siguiente">Siguiente</button>' : '<button class="primario" data-accion="crear">Guardar proyecto</button>'}</div>`;
}

function leerPasoGeneral() {
  const d = ST.borrador, v = id => document.getElementById(id);
  if (v('f-nombre')) {
    d.proyecto = v('f-nombre').value.trim();
    if (v('f-inicio').value) d.inicio = lunesDe(v('f-inicio').value);
    d.n_sem = Math.max(1, Math.min(104, Math.round(n(v('f-nsem').value) || 24)));
    d.cont_pct = (n(v('f-cont').value) || 0) / 100;
    d.umbral = n(v('f-umbral').value) || 0.95; d.umbral_act = n(v('f-umbral-act').value) || 0.9;
  }
  if (v('f-total')) d.total_control = n(v('f-total').value);
}

// =================================================================== proyecto: informe, semana, datos, exportar
function resultado(corte) {
  try { return calcularDatos(ST.datos, corte ? { corte } : {}); } catch (e) { return { error: e.message }; }
}

function pintarProyecto() {
  const d = ST.datos; if (!d) { ST.vista = 'inicio'; return pintar(); }
  const tabs = [['informe', 'Informe'], ...(ST.escribir ? [['semana', 'Actualizar semana']] : []), ['datos', 'Datos'], ['exportar', 'Exportar']];
  vista.innerHTML = `
  <div class="cabecera-vista"><div><p class="muted"><button class="link" data-accion="inicio">← Proyectos</button></p><h1>${esc(d.proyecto)}</h1></div>
    ${ST.escribir ? '<button data-accion="configurar">Configuración</button>' : ''}</div>
  <nav class="tabs">${tabs.map(([k, t]) => `<button class="${ST.tab === k ? 'activa' : ''}" data-accion="tab" data-tab="${k}">${t}</button>`).join('')}</nav>
  <div id="contenido-tab"></div>`;
  const c = document.getElementById('contenido-tab');
  if (ST.tab === 'informe') {
    const r = resultado(ST.corteVista);
    if (r.error) { c.innerHTML = `<div class="card error">${esc(r.error)}</div>`; return; }
    c.innerHTML = `<div class="barra-informe"><label>Semana de corte <input type="number" id="corte-vista" min="1" max="104" value="${r.D.corte}"></label>
      <span class="muted">Guardada: semana ${d.corte}</span></div><div class="wrap" id="app"></div>`;
    render(r.D, 'incrustado');
  }
  if (ST.tab === 'semana') pintarSemana(c);
  if (ST.tab === 'datos') pintarDatos(c);
  if (ST.tab === 'exportar') c.innerHTML = `<div class="card"><h2>Exportar</h2>
    <div class="lista-export">
      <div><b>Informe para el directorio</b><span class="muted">HTML autocontenido; se abre en cualquier navegador y se imprime a PDF.</span><button class="primario" data-accion="exp-informe">Descargar informe</button></div>
      <div><b>Excel del proyecto</b><span class="muted">Misma estructura que la plantilla: se abre con App_Control_Proyecto.html sin conexión y sirve de copia de seguridad en la carpeta de Drive.</span><button data-accion="exp-excel">Descargar Excel</button></div>
      <div><b>Datos para Power BI</b><span class="muted">Una hoja por tabla, compatible con las medidas DAX de la guía.</span><button data-accion="exp-powerbi">Descargar datos Power BI</button></div>
    </div></div>`;
}

function pintarSemana(c) {
  const d = ST.datos;
  if (!ST.semana) ST.semana = { corte: Math.max(1, Math.min(104, semanaDe(d.inicio, hoyISO()))), revit: null, albaranes: [], planillas: [] };
  const s = ST.semana;
  const manuales = d.partidas.filter(p => p.fuente === 'MANUAL');
  const sinPagar = (d.planillas || []).map((x, i) => ({ ...x, i })).filter(x => x.sem_pago == null);
  const ultimo = cod => { const ser = (d.manual || {})[cod] || []; for (let w = Math.min(s.corte, ser.length); w >= 1; w--) if (ser[w - 1] != null) return ser[w - 1]; return 0; };
  c.innerHTML = `
  <div class="card"><h2>1 · Semana de corte</h2><label class="control">Semana<input type="number" id="s-corte" min="1" max="104" value="${s.corte}"></label>
    <span class="muted">Semana ${s.corte}: del ${fecha(new Date(new Date(d.inicio + 'T00:00:00Z').getTime() + (s.corte - 1) * 6048e5).toISOString().slice(0, 10))} al domingo siguiente.</span></div>
  <div class="card"><h2>2 · Avance en Revit</h2>
    <p>Adjunta el archivo <b>Revit_Avance_….csv</b> que genera el botón <b>Exportar avance</b> de Revit (o una tabla de planificación exportada con los parámetros CC_*).</p>
    <label class="drop"><input type="file" id="s-revit" accept=".csv,.txt,.xlsx,.xls" hidden><b>Adjuntar exportación de Revit</b><span class="muted">${d.revit_info ? `Última: ${esc(d.revit_info.archivo)} (${fecha(d.revit_info.importado.slice(0, 10))}, ${d.revit_info.total} elementos)` : 'Aún no se ha importado ninguna.'}</span></label>
    ${s.revit ? `<div class="ok-box">Leído <b>${esc(s.revit.archivo)}</b>: ${s.revit.total} elementos con partida, ${s.revit.items.filter(x => x.estado === 'Instalado').reduce((a, x) => a + x.nElem, 0)} instalados. Se aplicará al guardar.</div>` : ''}</div>
  <div class="card"><h2>3 · Albaranes de entrega</h2>
    ${zonaIA('alb', 'Opcional: escribe lo entregado (p. ej. "albarán 1234: 2 compresores MYCOM 200").', '<button class="primario" data-accion="ia-albaranes">Leer con IA</button>')}
    ${tabla('t-alb', colsAlb(), s.albaranes)}</div>
  <div class="card"><h2>4 · Certificación / planilla del contratista</h2>
    ${zonaIA('pla', 'Opcional: escribe las líneas certificadas.', '<button class="primario" data-accion="ia-planilla">Leer con IA</button>')}
    ${tabla('t-pla', colsPla(), s.planillas)}</div>
  ${manuales.length ? `<div class="card"><h2>5 · Avance de lo que no está en el modelo</h2><div class="manuales">${manuales.map(p => `<label>${esc(p.cod)} · ${esc(p.desc)}
    <span class="pct"><input type="number" min="0" max="100" step="5" data-manual="${esc(p.cod)}" value="${Math.round(ultimo(p.cod) * 100)}"> %</span></label>`).join('')}</div></div>` : ''}
  ${sinPagar.length ? `<div class="card"><h2>6 · Pagos</h2><p class="muted">Marca lo que se ha pagado esta semana.</p><div class="pagos">${sinPagar.map(x => `<label><input type="checkbox" data-pago="${x.i}"> ${esc(x.n || 'Sin nº')} · S${x.sem} · ${esc(x.cod)} · ${eur(Math.round(x.q * x.pu * 100) / 100)}</label>`).join('')}</div></div>` : ''}
  ${(d.extras || []).length ? `<div class="card"><h2>7 · Decisiones sobre extras</h2>${d.extras.map((e, i) => `<label class="fila-extra">${esc(e.desc)} (${eur(e.imp)})<select data-extra="${i}">${['Sin decidir', 'Aprobado', 'Rechazado'].map(o => `<option ${o === e.estado ? 'selected' : ''}>${o}</option>`).join('')}</select></label>`).join('')}</div>` : ''}
  <div class="navegacion"><span></span><button class="primario" data-accion="guardar-semana">Guardar semana y ver informe</button></div>`;
}

function pintarDatos(c) {
  const d = ST.datos;
  const pl = (d.planillas || []), su = (d.suministros || []);
  c.innerHTML = `
  <div class="card"><h2>Albaranes registrados (${su.length})</h2>${su.length ? `<div class="scroll"><table><thead><tr><th>Semana</th><th>Partida</th><th class="n">Cantidad</th><th>Albarán</th>${ST.escribir ? '<th></th>' : ''}</tr></thead><tbody>
    ${su.map((x, i) => `<tr><td>S${x.sem}</td><td>${esc(x.cod)}</td><td class="n">${dec(x.q, 2)}</td><td>${esc(x.alb || '')}</td>${ST.escribir ? `<td><button class="link" data-accion="borrar-reg" data-lista="suministros" data-i="${i}">Borrar</button></td>` : ''}</tr>`).join('')}</tbody></table></div>` : '<p class="muted">Ninguno.</p>'}</div>
  <div class="card"><h2>Planillas registradas (${pl.length})</h2>${pl.length ? `<div class="scroll"><table><thead><tr><th>Nº</th><th>Semana</th><th>Partida</th><th class="n">Cantidad</th><th class="n">Precio</th><th class="n">Importe</th><th>Pago</th>${ST.escribir ? '<th></th>' : ''}</tr></thead><tbody>
    ${pl.map((x, i) => `<tr><td>${esc(x.n)}</td><td>S${x.sem}</td><td>${esc(x.cod)}</td><td class="n">${dec(x.q, 2)}</td><td class="n">${eur(x.pu, 2)}</td><td class="n">${eur(x.q * x.pu, 2)}</td><td>${x.sem_pago ? 'S' + x.sem_pago : '–'}</td>${ST.escribir ? `<td><button class="link" data-accion="borrar-reg" data-lista="planillas" data-i="${i}">Borrar</button></td>` : ''}</tr>`).join('')}</tbody></table></div>` : '<p class="muted">Ninguna.</p>'}</div>
  <div class="card"><h2>Presupuesto y medición</h2><div class="scroll"><table><thead><tr><th>Partida</th><th>Descripción</th><th>Actividad</th><th>Se mide por</th><th class="n">Cant.</th><th class="n">Importe</th></tr></thead><tbody>
    ${d.partidas.map(p => `<tr><td>${esc(p.cod)}</td><td>${esc(p.desc)}</td><td>${esc(p.act)}</td><td>${esc(p.fuente)}</td><td class="n">${p.cmed ?? ''} ${esc(p.ud)}</td><td class="n">${eur(p.bac)}</td></tr>`).join('')}</tbody></table></div>
    ${ST.escribir ? '<p><button data-accion="configurar">Editar configuración</button></p>' : ''}</div>`;
}

// =================================================================== eventos
const adjuntos = {};  // id de zona -> File[]
function mostrarAdjuntos(id) { const el = document.getElementById('lista-' + id); if (el) el.textContent = (adjuntos[id] || []).map(f => f.name).join(' · '); }

document.addEventListener('change', async ev => {
  const t = ev.target;
  if (t.id && t.id.startsWith('adj-')) { const id = t.id.slice(4); adjuntos[id] = [...(adjuntos[id] || []), ...t.files]; mostrarAdjuntos(id); t.value = ''; }
  if (t.id === 'importar' && t.files[0]) {  // el selector está fuera de #vista: un repintado no lo destruye
    try {
      const wb = XLSX.read(new Uint8Array(await t.files[0].arrayBuffer()), { type: 'array' });
      const datos = leerLibro(wb); calcularDatos(datos, {});
      const id = nuevoId(); if (await guardar(id, datos)) { ST.id = id; ST.datos = datos; ST.vista = 'proyecto'; ST.tab = 'informe'; pintar(); }
    } catch (e) { mensaje('No se ha podido importar: ' + e.message, 'error'); }
    t.value = '';
  }
  if (t.id === 'corte-vista') { ST.corteVista = Math.round(n(t.value) || 1); pintar(); }
  if (t.id === 's-corte') { ST.semana.corte = Math.max(1, Math.min(104, Math.round(n(t.value) || 1))); leerTablas(); pintar(); }
  if (t.id === 's-revit' && t.files[0]) {
    try { ST.semana.revit = await leerRevit(t.files[0]); if (!ST.semana.revit.total) mensaje('El archivo no tiene elementos con CC_COD_PARTIDA.', 'warn'); }
    catch (e) { mensaje('No se ha podido leer el archivo de Revit: ' + e.message, 'error'); }
    leerTablas(); pintar();
  }
});
document.addEventListener('dragover', e => { const z = e.target.closest && e.target.closest('[data-drop]'); if (z) { e.preventDefault(); z.classList.add('sobre'); } });
document.addEventListener('dragleave', e => { const z = e.target.closest && e.target.closest('[data-drop]'); if (z) z.classList.remove('sobre'); });
document.addEventListener('drop', e => { const z = e.target.closest && e.target.closest('[data-drop]'); if (!z) return; e.preventDefault(); z.classList.remove('sobre');
  const id = z.dataset.drop; adjuntos[id] = [...(adjuntos[id] || []), ...e.dataTransfer.files]; mostrarAdjuntos(id); });

document.addEventListener('click', async ev => {
  const b = ev.target.closest('[data-accion]'); if (!b || b.tagName === 'INPUT') return;
  const a = b.dataset.accion;
  if (a === 'parar-ia') { ctlIA && ctlIA.abort(); return; }
  if (ST.vista === 'asistente' && a !== 'cancelar') { leerPasoGeneral(); leerTablas(); }
  if (ST.vista === 'proyecto' && ST.tab === 'semana') leerTablas();
  const d = ST.borrador;

  if (a === 'nuevo') { ST.borrador = nuevoBorrador(); ST.paso = 0; ST.id = null; ST.vista = 'asistente'; ST.editando = true; return pintar(); }
  if (a === 'abrir') { const p = ST.proyectos.find(x => x.id === b.dataset.id); ST.id = p.id; ST.datos = clon(p.datos); ST.vista = 'proyecto'; ST.tab = 'informe'; ST.corteVista = null; ST.semana = null; ST.editando = false; return pintar(); }
  if (a === 'inicio') { ST.vista = 'inicio'; ST.id = null; ST.editando = false; return pintar(); }
  if (a === 'configurar') { ST.borrador = clon(ST.datos); ST.paso = 0; ST.vista = 'asistente'; ST.editando = true; return pintar(); }
  if (a === 'cancelar') { ST.editando = false; ST.vista = ST.id ? 'proyecto' : 'inicio'; return pintar(); }
  if (a === 'anterior') { ST.paso--; return pintar(); }
  if (a === 'siguiente') {
    if (ST.paso === 0 && !d.proyecto) return mensaje('Escribe el nombre del proyecto.', 'warn');
    if (ST.paso === 1 && !d.partidas.length) return mensaje('Añade al menos una partida (adjunta la oferta o escríbelas).', 'warn');
    if (ST.paso === 2 && !d.actividades.length) return mensaje('Añade al menos una actividad.', 'warn');
    ST.paso++; return pintar();
  }
  if (a === 'crear') {
    try { calcularDatos(d, {}); } catch (e) { return mensaje(e.message, 'error'); }
    const id = ST.id || nuevoId();
    if (await guardar(id, d)) { ST.id = id; ST.datos = clon(d); ST.vista = 'proyecto'; ST.tab = 'informe'; ST.editando = false; pintar(); }
    return;
  }
  if (a === 'anadir-fila') { const def = TABLAS[b.dataset.tabla]; def.lista(ST.vista === 'asistente' ? d : ST.datos).push(def.nueva(ST.vista === 'asistente' ? d : ST.datos)); return pintar(); }
  if (a === 'borrar-fila') { TABLAS[b.dataset.tabla].lista(ST.vista === 'asistente' ? d : ST.datos).splice(+b.dataset.i, 1); return pintar(); }
  if (a === 'tab') { ST.tab = b.dataset.tab; return pintar(); }

  // ---------- IA en el asistente
  if (a === 'ia-presupuesto') {
    const r = await pedirIA(b, PROMPTS.presupuesto, adjuntos.pres, document.getElementById('txt-pres').value);
    if (!r) return;
    const usados = new Set();
    d.partidas = (r.partidas || []).filter(x => n(x.importe) != null).map((x, i) => {
      let cod = String(x.cod || '').trim() || 'P' + String(i + 1).padStart(2, '0');
      while (usados.has(cod)) cod += '_'; usados.add(cod);
      return { cod, cap: String(x.capitulo || ''), desc: String(x.descripcion || ''), cant: n(x.cantidad), ud: String(x.unidad || 'ud'), bac: n(x.importe),
               act: '', fuente: 'MANUAL', perfil: 'LINEAL', cmed: n(x.cantidad), sem_ent: null };
    });
    d.extras = (r.extras || []).filter(x => n(x.importe) != null).map((x, i) => ({ cod: 'X' + String(i + 1).padStart(2, '0'), desc: String(x.descripcion || ''), imp: n(x.importe), estado: 'Sin decidir', limite: null, motivo: '' }));
    if (n(r.total_documento) != null && d.total_control == null) d.total_control = n(r.total_documento);
    pintar(); if (r.notas) mensaje('IA: ' + r.notas, 'info'); return;
  }
  if (a === 'ia-cronograma' || a === 'ia-proponer-crono') {
    const r = a === 'ia-cronograma' ? await pedirIA(b, PROMPTS.cronograma(d.inicio), adjuntos.crono, document.getElementById('txt-crono').value)
                                    : await pedirIA(b, PROMPTS.proponerCronograma(d.partidas, d.n_sem), [], document.getElementById('txt-crono').value || 'Propón el cronograma.');
    if (!r) return;
    d.actividades = (r.actividades || []).map(x => ({ id: String(x.id || '').trim(), nombre: String(x.nombre || ''), si: Math.max(1, Math.round(n(x.sem_inicio) || 1)), sf: Math.max(1, Math.round(n(x.sem_fin) || 1)), pred: String(x.predecesora || '') })).filter(x => x.id);
    const fin = Math.max(0, ...d.actividades.map(x => x.sf)); if (fin > d.n_sem) d.n_sem = Math.min(104, fin);
    pintar(); if (r.notas) mensaje('IA: ' + r.notas, 'info'); return;
  }
  if (a === 'ia-medicion') {
    if (!d.actividades.length) return mensaje('Primero define las actividades (paso 3).', 'warn');
    const r = await pedirIA(b, PROMPTS.medicion(d.partidas, d.actividades), [], 'Propón la medición.');
    if (!r) return;
    for (const x of r.asignaciones || []) {
      const p = d.partidas.find(y => y.cod === x.cod); if (!p) continue;
      if (d.actividades.some(z => z.id === x.act)) p.act = x.act;
      if (['ALBARAN', 'REVIT', 'MANUAL', 'DERIVADO'].includes(x.fuente)) p.fuente = x.fuente;
      p.perfil = x.perfil === 'ENTREGA' ? 'ENTREGA' : 'LINEAL'; if (x.ud) p.ud = String(x.ud); p.cmed = n(x.cmed);
      if (n(x.horas) != null) d.horas[p.cod] = n(x.horas);
    }
    pintar(); if ((r.dudas || []).length) mensaje('IA: ' + r.dudas.join(' · '), 'info'); return;
  }
  // ---------- semana
  if (a === 'ia-albaranes') {
    const r = await pedirIA(b, PROMPTS.albaranes(ST.datos.partidas, ST.semana.corte), adjuntos.alb, document.getElementById('txt-alb').value);
    if (!r) return;
    ST.semana.albaranes.push(...(r.lineas || []).map(x => ({ sem: x.fecha ? Math.max(1, semanaDe(ST.datos.inicio, x.fecha)) : ST.semana.corte, cod: String(x.cod || ''), q: n(x.cantidad), alb: String(x.albaran || ''), texto: String(x.texto_original || '') })));
    adjuntos.alb = []; pintar(); if (r.notas) mensaje('IA: ' + r.notas, 'info'); return;
  }
  if (a === 'ia-planilla') {
    const r = await pedirIA(b, PROMPTS.planilla(ST.datos.partidas), adjuntos.pla, document.getElementById('txt-pla').value);
    if (!r) return;
    ST.semana.planillas.push(...(r.lineas || []).map(x => ({ n: String(r.numero || ''), sem: ST.semana.corte, cod: String(x.cod || ''), q: n(x.cantidad),
      pu: n(x.precio) ?? (n(x.importe) != null && n(x.cantidad) ? n(x.importe) / n(x.cantidad) : null), obs: String(x.observacion || '') })));
    adjuntos.pla = []; pintar(); if (r.notas) mensaje('IA: ' + r.notas, 'info'); return;
  }
  if (a === 'guardar-semana') {
    const s = ST.semana, dd = clon(ST.datos);
    const malAlb = s.albaranes.filter(x => !x.cod || x.q == null), malPla = s.planillas.filter(x => !x.cod || x.q == null || x.pu == null);
    if (malAlb.length || malPla.length) return mensaje('Completa partida y cantidad (y precio en planillas) de todas las líneas, o quítalas.', 'warn');
    dd.corte = s.corte;
    if (s.revit) { dd.revit = s.revit.items; dd.revit_info = { archivo: s.revit.archivo, importado: s.revit.importado, total: s.revit.total }; }
    dd.suministros = [...(dd.suministros || []), ...s.albaranes.map(x => ({ sem: Math.round(x.sem || s.corte), cod: x.cod, q: x.q, alb: x.alb }))];
    dd.planillas = [...(dd.planillas || []), ...s.planillas.map(x => ({ n: x.n, sem: Math.round(x.sem || s.corte), cod: x.cod, q: x.q, pu: x.pu, obs: x.obs, sem_pago: null }))];
    dd.manual = dd.manual || {};
    for (const el of document.querySelectorAll('[data-manual]')) {
      const cod = el.dataset.manual, v = Math.min(100, Math.max(0, n(el.value) || 0)) / 100;
      const ser = dd.manual[cod] = dd.manual[cod] || [];
      let previo = 0; for (let w = 1; w < s.corte; w++) { if (ser[w - 1] == null) ser[w - 1] = previo; previo = ser[w - 1]; }  // rellena semanas sin dato
      ser[s.corte - 1] = v;
    }
    for (const el of document.querySelectorAll('[data-pago]')) if (el.checked) dd.planillas[+el.dataset.pago].sem_pago = s.corte;
    for (const el of document.querySelectorAll('[data-extra]')) dd.extras[+el.dataset.extra].estado = el.value;
    if (await guardar(ST.id, dd)) { ST.datos = dd; ST.semana = null; ST.tab = 'informe'; ST.corteVista = null; adjuntos.alb = []; adjuntos.pla = []; pintar(); }
    return;
  }
  if (a === 'borrar-reg') {
    const dd = clon(ST.datos); dd[b.dataset.lista].splice(+b.dataset.i, 1);
    if (await guardar(ST.id, dd)) { ST.datos = dd; pintar(); } return;
  }
  // ---------- exportar
  if (a === 'exp-informe') { const r = resultado(ST.corteVista); if (r.error) return mensaje(r.error, 'error'); return guardarArchivo(`Informe_${slug(r.D.proyecto)}_S${String(r.D.corte).padStart(2, '0')}.html`, await informeHTML(r.D)); }
  if (a === 'exp-excel') return guardarArchivo(`Control_${slug(ST.datos.proyecto)}.xlsx`, xlsxBlob(excelProyecto(ST.datos)));
  if (a === 'exp-powerbi') {
    const r = resultado(ST.corteVista); if (r.error) return mensaje(r.error, 'error');
    const wb = XLSX.utils.book_new(); for (const [k, filas] of Object.entries(r.tablas)) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(filas), k);
    return guardarArchivo(`PowerBI_${slug(ST.datos.proyecto)}.xlsx`, xlsxBlob(wb));
  }
});

window.__app = { ST, leerRevit, excelProyecto, informeHTML, pintar };  // para pruebas
if (window.pdfjsLib) pdfjsLib.GlobalWorkerOptions.workerSrc = './pdf.worker.min.js';
iniciar();
}
