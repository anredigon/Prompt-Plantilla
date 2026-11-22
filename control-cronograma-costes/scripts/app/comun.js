/* Utilidades comunes de la app de escritorio: números, tablas editables, Revit y Excel del proyecto. */
const n = v => { if (typeof v === 'number') return isFinite(v) ? v : null; const t = String(v ?? '').trim().replace(/\s|€/g, ''); if (!t) return null;
  const s = t.includes(',') ? t.replace(/\./g, '').replace(',', '.') : t; const x = Number(s); return isFinite(x) ? x : null; };
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

