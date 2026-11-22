/* Motor de cálculo de valor ganado (EVM). Lo usan la app sin conexión (Excel) y la app de claude.ai.
 *   leerLibro(wb)                -> datos   (lee el Excel de la plantilla con SheetJS)
 *   calcularDatos(datos, {corte}) -> { D, tablas, avisos }
 *   calcular(wb, {corte})        -> calcularDatos(leerLibro(wb), {corte})
 * Estructura de "datos" (JSON puro, se guarda tal cual en la base de datos de la app):
 *   { proyecto, inicio:'AAAA-MM-DD', n_sem, corte, cont_pct, umbral, umbral_act,
 *     actividades:[{id,nombre,si,sf,pred}], partidas:[{cod,cap,desc,cant,bac,act,fuente,ud,cmed,perfil,sem_ent}],
 *     revit:[{cod,q,estado,sem}], manual:{cod:[%acum por semana]}, suministros:[{sem,cod,q,alb}], horas:{cod:h/ud},
 *     planillas:[{n,sem,cod,q,pu,obs,sem_pago}], extras:[{cod,desc,imp,estado,limite,motivo}],
 *     contingencia:[{sem,concepto,imp,estado}], avisos:[] }
 * Reglas (validadas por dirección): ver README, sección "Reglas de cálculo".
 */
function calcular(wb, opciones) { return calcularDatos(leerLibro(wb), opciones); }

function leerLibro(wb) {
  const avisos = [];
  const HOJAS = ['Parametros', 'Cronograma', 'Oferta', 'Revit_Export', 'Suministros', 'Horas_Montaje',
                 'Avance_Manual', 'Planillas', 'Extras', 'Contingencia'];
  const faltan = HOJAS.filter(h => !wb.Sheets[h]);
  if (faltan.length) throw new Error('El Excel no tiene la estructura de la plantilla. Faltan las hojas: ' + faltan.join(', '));

  const filasDe = h => XLSX.utils.sheet_to_json(wb.Sheets[h], { header: 1, raw: true, defval: null }).slice(1);
  const valor = (h, ref) => { const c = wb.Sheets[h][ref]; return c ? c.v : null; };
  let numerosTexto = 0;
  // Números guardados como texto (pegados de otro programa). Criterio explícito:
  // "1.234,5" y "12,5" -> formato español; "1.234" o "1.234.567" -> miles; "12.5" -> punto decimal.
  const num = v => {
    if (typeof v === 'number' && isFinite(v)) return v;
    if (typeof v !== 'string') return null;
    let t = v.trim().replace(/\s|€|%$/g, '');
    if (t === '' || !/^-?[\d.,]+$/.test(t)) return null;
    if (t.includes(',')) t = t.replace(/\./g, '').replace(',', '.');
    else if (/^-?\d{1,3}(\.\d{3})+$/.test(t)) t = t.replace(/\./g, '');
    const n = Number(t);
    if (!isFinite(n)) return null;
    numerosTexto++;
    return v.trim().endsWith('%') ? n / 100 : n;
  };
  const txt = v => (v == null ? '' : String(v).trim());
  const r2 = x => Math.round(x * 100) / 100;
  const div = (a, b) => (b ? a / b : null);
  const MS_DIA = 864e5;
  const sistema1904 = !!(wb.Workbook && wb.Workbook.WBProps && wb.Workbook.WBProps.date1904);
  const serialAFecha = v => {
    if (typeof v === 'number') return new Date(Math.round((v + (sistema1904 ? 1462 : 0) - 25569) * MS_DIA));
    if (v instanceof Date) return new Date(Date.UTC(v.getFullYear(), v.getMonth(), v.getDate()));
    return null;
  };
  const iso = d => d.toISOString().slice(0, 10);
  const masDias = (d, n) => new Date(d.getTime() + n * MS_DIA);

  // ---------------------------------------------------------------- parámetros
  const proyecto = txt(valor('Parametros', 'B3')) || 'Proyecto sin nombre';
  const inicio = serialAFecha(valor('Parametros', 'B4'));
  if (!inicio) throw new Error('Parametros!B4 (fecha de inicio) no es una fecha válida.');
  const n_sem = Math.round(num(valor('Parametros', 'B5')) || 0);
  if (n_sem < 1 || n_sem > 104) throw new Error('Parametros!B5 (nº de semanas) debe estar entre 1 y 104.');
  const corte = Math.round(num(valor('Parametros', 'B6')) || 1);
  const cont_pct = num(valor('Parametros', 'B8')) || 0;
  const umbral = num(valor('Parametros', 'B9')) || 0.95;
  const umbral_act = num(valor('Parametros', 'B10')) || 0.9;

  // ---------------------------------------------------------------- lectura
  const actividades = [];
  for (const r of filasDe('Cronograma')) {
    if (!txt(r[0])) continue;
    actividades.push({ id: txt(r[0]), nombre: txt(r[1]), si: num(r[2]), sf: num(r[3]), pred: txt(r[7]) });
  }

  const partidas = [];
  for (const r of filasDe('Oferta')) {
    const cod = txt(r[0]);
    if (!cod || num(r[4]) == null || !txt(r[6])) continue;
    partidas.push({ cod, cap: txt(r[1]), desc: txt(r[2]), cant: num(r[3]), bac: num(r[4]), act: txt(r[6]),
                    fuente: txt(r[7]).toUpperCase(), ud: txt(r[8]), cmed: num(r[9]),
                    perfil: txt(r[22]).toUpperCase() === 'ENTREGA' ? 'ENTREGA' : 'LINEAL',
                    sem_ent: num(r[23]) != null ? Math.round(num(r[23])) : null });
  }

  const revit = [];
  for (const r of filasDe('Revit_Export')) {
    const cod = txt(r[4]);
    if (!cod) continue;
    const sem = num(r[9]);
    revit.push({ cod, q: num(r[6]) || 0, estado: txt(r[8]), sem: sem == null ? null : Math.round(sem) });
  }

  const manual = {};
  for (const r of filasDe('Avance_Manual')) {
    if (!txt(r[0])) continue;
    manual[txt(r[0])] = r.slice(2, 2 + 104).map(v => num(v));
  }
  const suministros = [];
  for (const r of filasDe('Suministros')) {
    if (num(r[0]) == null || !txt(r[1])) continue;
    suministros.push({ sem: Math.round(num(r[0])), cod: txt(r[1]), q: num(r[2]) || 0, alb: txt(r[3]) });
  }
  const horas = {};
  for (const r of filasDe('Horas_Montaje')) if (txt(r[0]) && num(r[2]) != null) horas[txt(r[0])] = num(r[2]);

  const planillas = [];
  for (const r of filasDe('Planillas')) {
    if (num(r[1]) == null || !txt(r[2]) || num(r[3]) == null || num(r[4]) == null) continue;
    const pago = num(r[7]);
    planillas.push({ n: txt(r[0]), sem: Math.round(num(r[1])), cod: txt(r[2]), q: num(r[3]), pu: num(r[4]),
                     obs: txt(r[6]), sem_pago: pago == null ? null : Math.round(pago) });
  }

  const extras = [];
  for (const r of filasDe('Extras')) {
    if (!txt(r[0]) || num(r[2]) == null) continue;
    const lim = num(r[4]);
    extras.push({ cod: txt(r[0]), desc: txt(r[1]), imp: num(r[2]), estado: txt(r[3]) || 'Sin decidir',
                  limite: lim == null ? null : Math.round(lim), motivo: txt(r[5]) });
  }
  const contingencia = [];
  for (const r of filasDe('Contingencia')) {
    if (num(r[0]) == null || num(r[2]) == null) continue;
    contingencia.push({ sem: Math.round(num(r[0])), concepto: txt(r[1]), imp: num(r[2]), estado: txt(r[3]) });
  }
  if (numerosTexto) avisos.push(`Hay ${numerosTexto} números guardados como texto en el Excel; se han interpretado (coma = decimal). Revísalos o conviértelos a número.`);
  return { proyecto, inicio: iso(inicio), n_sem, corte, cont_pct, umbral, umbral_act, actividades, partidas, revit,
           manual, suministros, horas, planillas, extras, contingencia, avisos };
}

function calcularDatos(datos, opciones) {
  opciones = opciones || {};
  const avisos = [...(datos.avisos || [])];
  const r2 = x => Math.round(x * 100) / 100;
  const div = (a, b) => (b ? a / b : null);
  const MS_DIA = 864e5;
  const iso = d => d.toISOString().slice(0, 10);
  const masDias = (d, n) => new Date(d.getTime() + n * MS_DIA);
  const esNum = v => typeof v === 'number' && isFinite(v);

  // ---------------------------------------------------------------- parámetros
  const proyecto = String(datos.proyecto || '').trim() || 'Proyecto sin nombre';
  const inicio = /^\d{4}-\d{2}-\d{2}$/.test(datos.inicio || '') ? new Date(datos.inicio + 'T00:00:00Z') : null;
  if (!inicio || isNaN(inicio)) throw new Error('La fecha de inicio del proyecto no es válida.');
  const n_sem = Math.round(datos.n_sem || 0);
  if (n_sem < 1 || n_sem > 104) throw new Error('El nº de semanas del cronograma debe estar entre 1 y 104.');
  let corte = Math.round(esNum(opciones.corte) ? opciones.corte : (datos.corte || 1));
  if (corte > 104) avisos.push(`La semana de corte (${corte}) supera el máximo de 104; se usa 104.`);
  corte = Math.min(Math.max(corte, 1), 104);
  // Horizonte del informe: si la obra sigue abierta después del plan, se informa más allá de n_sem con PV plano
  const horizonte = Math.max(n_sem, corte);
  if (corte > n_sem) avisos.push(`La semana de corte (${corte}) es posterior al fin planificado (semana ${n_sem}): la obra va fuera de plazo.`);
  const cont_pct = esNum(datos.cont_pct) ? datos.cont_pct : 0;
  const umbral = esNum(datos.umbral) && datos.umbral > 0 ? datos.umbral : 0.95;
  const umbral_act = esNum(datos.umbral_act) && datos.umbral_act > 0 ? datos.umbral_act : 0.9;

  // ---------------------------------------------------------------- validación
  const actividades = [];
  for (const a of datos.actividades || []) {
    const id = String(a.id || '').trim();
    if (!id) continue;
    if (actividades.some(x => x.id === id)) { avisos.push(`Cronograma: la actividad ${id} está repetida; se usa solo la primera.`); continue; }
    if (!esNum(a.si) || !esNum(a.sf) || a.sf < a.si) { avisos.push(`Cronograma: la actividad ${id} no tiene semanas de inicio/fin válidas; se ignora.`); continue; }
    actividades.push({ id, nombre: String(a.nombre || ''), si: Math.round(a.si), sf: Math.round(a.sf), pred: String(a.pred || '') });
  }
  const actPorId = Object.fromEntries(actividades.map(a => [a.id, a]));
  const FUENTES = ['ALBARAN', 'REVIT', 'MANUAL', 'DERIVADO'];
  const partidas = [];
  for (const p of datos.partidas || []) {
    const cod = String(p.cod || '').trim();
    if (!cod || !esNum(p.bac)) continue;
    if (partidas.some(x => x.cod === cod)) { avisos.push(`Oferta: el código de partida ${cod} está repetido; se usa solo la primera fila. Corrige el código para que cuente.`); continue; }
    const a = actPorId[String(p.act || '').trim()];
    if (!a) { avisos.push(`Oferta: la partida ${cod} apunta a la actividad "${p.act || ''}", que no está en el Cronograma; se ignora.`); continue; }
    const fuente = String(p.fuente || '').toUpperCase();
    if (!FUENTES.includes(fuente)) { avisos.push(`Oferta: la partida ${cod} tiene una fuente de avance no válida ("${p.fuente || ''}"); se ignora.`); continue; }
    if ((fuente === 'REVIT' || fuente === 'ALBARAN') && !(p.cmed > 0)) { avisos.push(`Oferta: la partida ${cod} necesita una cantidad de medición mayor que 0; se ignora.`); continue; }
    partidas.push({ cod, cap: String(p.cap || ''), desc: String(p.desc || ''), bac: p.bac, act: a.id, fuente, ud: String(p.ud || ''),
                    cmed: p.cmed > 0 ? p.cmed : 1, perfil: p.perfil === 'ENTREGA' ? 'ENTREGA' : 'LINEAL',
                    sem_ent: esNum(p.sem_ent) ? Math.round(p.sem_ent) : a.si });
  }
  if (!partidas.length) throw new Error('El presupuesto no tiene ninguna partida válida (código, importe y actividad).');
  const codigos = new Set(partidas.map(p => p.cod));

  const revit = [];
  const estadosRaros = new Set();
  let sinSemana = 0;
  for (const e of datos.revit || []) {
    const cod = String(e.cod || '').trim();
    if (!cod) continue;
    const est = String(e.estado || '').trim().toLowerCase();  // igual que SUMIFS de Excel: sin distinguir mayúsculas
    if (est && est !== 'instalado' && est !== 'pendiente') estadosRaros.add(e.estado);
    const sem = esNum(e.sem) ? Math.round(e.sem) : null;
    if (est === 'instalado' && sem == null) sinSemana++;
    revit.push({ cod, q: esNum(e.q) ? e.q : 0, estado: est === 'instalado' ? 'Instalado' : 'Pendiente', sem });
  }
  if (estadosRaros.size) avisos.push(`Revit: CC_ESTADO con valores no reconocidos (cuentan como Pendiente): ${[...estadosRaros].slice(0, 5).join(', ')}.`);
  if (sinSemana) avisos.push(`Revit: ${sinSemana} elementos "Instalado" sin CC_SEMANA_INSTALADO no suman avance.`);
  const huerfanos = new Set(revit.filter(e => !codigos.has(e.cod)).map(e => e.cod));
  if (huerfanos.size) avisos.push(`Revit: hay elementos con CC_COD_PARTIDA que no existe en el presupuesto: ${[...huerfanos].slice(0, 8).join(', ')}.`);

  const manual = {};
  const fueraRango = new Set();
  for (const [cod, serieM] of Object.entries(datos.manual || {})) {
    manual[cod] = (serieM || []).map(v => {
      if (!esNum(v)) return null;
      if (v < 0 || v > 1) fueraRango.add(cod);
      return Math.min(1, Math.max(0, v));
    });
  }
  if (fueraRango.size) avisos.push(`Avance manual: hay % fuera de 0–100 % en ${[...fueraRango].join(', ')}; se limitan a 0–100 %.`);
  const suministros = (datos.suministros || []).filter(x => esNum(x.sem) && x.cod).map(x => ({ sem: Math.round(x.sem), cod: String(x.cod), q: esNum(x.q) ? x.q : 0 }));
  const horas = {};
  for (const [cod, h] of Object.entries(datos.horas || {})) if (esNum(h)) horas[cod] = h;
  const planillas = (datos.planillas || []).filter(x => esNum(x.sem) && x.cod && esNum(x.q) && esNum(x.pu)).map(x => ({
    n: String(x.n || ''), sem: Math.round(x.sem), cod: String(x.cod), q: x.q, pu: x.pu, imp: r2(x.q * x.pu), obs: String(x.obs || ''),
    sem_pago: esNum(x.sem_pago) ? Math.round(x.sem_pago) : null }));
  const sinPartida = new Set(planillas.filter(x => !codigos.has(x.cod)).map(x => x.cod));
  if (sinPartida.size) avisos.push(`Planillas: hay líneas con partidas que no existen en el presupuesto (no cuentan en el coste): ${[...sinPartida].join(', ')}.`);
  const extras = (datos.extras || []).filter(e => e.cod && esNum(e.imp)).map(e => ({ cod: String(e.cod), desc: String(e.desc || ''), imp: e.imp,
    estado: e.estado || 'Sin decidir', limite: esNum(e.limite) ? Math.round(e.limite) : null, motivo: String(e.motivo || '') }));
  const contingencia = (datos.contingencia || []).filter(c => esNum(c.sem) && esNum(c.imp)).map(c => ({ sem: Math.round(c.sem), concepto: String(c.concepto || ''), imp: c.imp, estado: String(c.estado || '') }));

  // ---------------------------------------------------------------- cálculo semanal
  const SEMANAS = Array.from({ length: horizonte }, (_, i) => i + 1);
  const pvAcum = (p, w) => {
    if (p.perfil === 'ENTREGA') return w >= p.sem_ent ? p.bac : 0;
    const a = actPorId[p.act];
    return p.bac * Math.min(1, Math.max(0, (w - a.si + 1) / (a.sf - a.si + 1)));
  };
  const cantRevit = (p, w) => revit.reduce((s, e) => s + (e.cod === p.cod && e.estado === 'Instalado' && e.sem != null && e.sem <= w ? e.q : 0), 0);
  const cantEntregada = (p, w) => suministros.reduce((s, x) => s + (x.cod === p.cod && x.sem <= w ? x.q : 0), 0);
  const cantMedida = (p, w) => (p.fuente === 'ALBARAN' ? cantEntregada(p, w) : cantRevit(p, w));
  const horasTot = revit.reduce((s, e) => s + (horas[e.cod] || 0) * e.q, 0);
  const fracHoras = w => (horasTot ? revit.reduce((s, e) => s + (e.estado === 'Instalado' && e.sem != null && e.sem <= w ? (horas[e.cod] || 0) * e.q : 0), 0) / horasTot : 0);
  if (partidas.some(p => p.fuente === 'DERIVADO') && !horasTot) avisos.push('Hay partidas DERIVADO (montaje) pero no hay horas de montaje para elementos del modelo de Revit (o aún no se ha importado Revit): su avance será 0 hasta entonces.');
  const pctAvance = (p, w) => {
    if (p.fuente === 'REVIT' || p.fuente === 'ALBARAN') return Math.min(1, cantMedida(p, w) / p.cmed);
    if (p.fuente === 'MANUAL') { const s = manual[p.cod] || []; return s[w - 1] || 0; }
    return fracHoras(w);
  };

  const serie = {};
  for (const p of partidas) for (const w of SEMANAS) {
    const d = { pv: pvAcum(p, w) };
    if (w <= corte) {
      d.ev = p.bac * pctAvance(p, w);
      d.ac = planillas.reduce((s, x) => s + (x.cod === p.cod && x.sem <= w ? x.imp : 0), 0);
    }
    serie[p.cod + '|' + w] = d;
  }
  const tot = (campo, w, cods) => partidas.reduce((s, p) => s + ((!cods || cods.has(p.cod)) ? (serie[p.cod + '|' + w][campo] || 0) : 0), 0);

  const BAC = partidas.reduce((s, p) => s + p.bac, 0);
  const PV = tot('pv', corte), EV = tot('ev', corte), AC = tot('ac', corte);
  const SPI = div(EV, PV), CPI = div(EV, AC);
  const EAC = CPI ? BAC / CPI : AC + (BAC - EV);
  const curva = SEMANAS.map(w => ({ sem: w, pv: r2(tot('pv', w)), ev: w <= corte ? r2(tot('ev', w)) : null, ac: w <= corte ? r2(tot('ac', w)) : null }));

  // Earned schedule: el momento MÁS TEMPRANO en que el plan alcanzaba el valor ganado hoy.
  // Así una meseta del plan (semanas sin nada planificado) nunca se acredita como avance:
  // con EV = 0 el ES es 0. C = última semana con PV acumulado estrictamente menor que EV.
  const pvs = [0, ...curva.slice(0, n_sem).map(c => c.pv)];
  let ES;
  if (EV <= 1e-6) ES = 0;
  else if (EV >= pvs[pvs.length - 1] - 1e-6) ES = n_sem;
  else {
    let C = 0;
    for (let t = 0; t < pvs.length; t++) if (pvs[t] < EV - 1e-6) C = t;
    ES = C + (EV - pvs[C]) / (pvs[C + 1] - pvs[C]);
  }
  let SPIt = PV > 0 ? div(ES, corte) : null;
  let dur_prev = SPIt ? n_sem / SPIt : null;
  const terminado = EV >= BAC - 0.01;
  if (terminado) {
    const finReal = Math.min(...curva.filter(c => c.ev != null && c.ev >= BAC - 0.01).map(c => c.sem));
    dur_prev = finReal; SPIt = n_sem / finReal;
  }
  const retraso = terminado ? 0 : Math.max(0, corte - ES);
  const finSem = w => masDias(inicio, 7 * w - 3);
  const fin_plan = finSem(n_sem);
  const fin_prev = dur_prev ? masDias(inicio, Math.round(7 * dur_prev) - 3) : null;
  const tcpi_inalcanzable = BAC - EV > 0.01 && AC >= BAC;
  const TCPI = (tcpi_inalcanzable || BAC - EV <= 0.01) ? null : div(BAC - EV, BAC - AC);

  const ext_pend = extras.filter(e => e.estado === 'Sin decidir').reduce((s, e) => s + e.imp, 0);
  const cont_ini = cont_pct * BAC;
  const cont_cons = contingencia.filter(c => c.estado === 'Aprobado').reduce((s, c) => s + c.imp, 0);
  const cont_prop = contingencia.filter(c => c.estado === 'Propuesto').reduce((s, c) => s + c.imp, 0);
  const cont_disp = cont_ini - cont_cons - cont_prop;
  const PAGADO = planillas.filter(x => codigos.has(x.cod) && x.sem <= corte && x.sem_pago != null && x.sem_pago <= corte).reduce((s, x) => s + x.imp, 0);
  const pagoAntes = planillas.filter(x => x.sem_pago != null && x.sem_pago < x.sem);
  if (pagoAntes.length) avisos.push(`Planillas: ${pagoAntes.length} líneas tienen semana de pago anterior a su semana de certificación (${[...new Set(pagoAntes.map(x => x.n))].slice(0, 5).join(', ')}).`);

  // ---------------------------------------------------------------- por actividad
  const ORDEN = { critical: 0, serious: 1, warning: 2, good: 3 };
  const nivel = (v, u) => { u = u == null ? umbral_act : u; if (v == null || v >= u) return 'good'; return v >= 0.85 ? 'warning' : (v >= 0.7 ? 'serious' : 'critical'); };
  const filasAct = actividades.map(a => {
    const cods = new Set(partidas.filter(p => p.act === a.id).map(p => p.cod));
    const bac = partidas.filter(p => cods.has(p.cod)).reduce((s, p) => s + p.bac, 0);
    const pv = tot('pv', corte, cods), ev = tot('ev', corte, cods), ac = tot('ac', corte, cods);
    const spi = pv ? ev / pv : null, cpi = ac ? ev / ac : null;
    let estado;
    if (pv === 0 && ev === 0 && ac === 0) estado = 'sin_iniciar';
    else estado = [nivel(spi), nivel(cpi)].sort((x, y) => ORDEN[x] - ORDEN[y])[0];
    return { id: a.id, nombre: a.nombre, si: a.si, sf: a.sf, bac: r2(bac), pv: r2(pv), ev: r2(ev), ac: r2(ac),
             pct_plan: bac ? pv / bac : 0, pct_real: bac ? ev / bac : 0, spi, cpi, estado };
  });

  // ---------------------------------------------------------------- alertas
  const miles = x => String(Math.round(Math.abs(x))).replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  const eur = x => (x < 0 ? '-' : '') + miles(x) + ' €';
  const dec2 = x => x.toFixed(2).replace('.', ',');
  const num2 = x => { const [e, d] = Math.abs(x).toFixed(2).split('.'); return (x < 0 ? '-' : '') + e.replace(/\B(?=(\d{3})+(?!\d))/g, '.') + ',' + d; };
  const pct0 = x => Math.round(x * 100) + '%';
  const alertas = [];
  for (const e of extras) {
    if (e.estado !== 'Sin decidir') continue;
    if (e.limite == null) { alertas.push({ nivel: 'warning', tipo: 'Decisión', texto: `${e.desc} (${eur(e.imp)}) sin decidir. ${e.motivo}`.trim() }); continue; }
    const margen = e.limite - corte;
    alertas.push({ nivel: margen <= 2 ? 'critical' : (margen <= 6 ? 'serious' : 'warning'), tipo: 'Decisión',
                   texto: `${e.desc} (${eur(e.imp)}) sin decidir. Límite: semana ${e.limite} (quedan ${margen} sem.). ${e.motivo}`.trim() });
  }
  for (const f of filasAct) {
    if (f.pv > 0 && f.pct_plan >= 0.999 && f.pct_real < 0.05)
      alertas.push({ nivel: 'critical', tipo: 'Plazo', texto: `${f.nombre}: debía estar terminada en la semana ${f.sf} y el avance es ${pct0(f.pct_real)}.` });
    else if (f.spi != null && ['serious', 'critical'].includes(nivel(f.spi)))
      alertas.push({ nivel: nivel(f.spi), tipo: 'Plazo', texto: `${f.nombre}: avance ${pct0(f.pct_real)} frente a ${pct0(f.pct_plan)} planificado (SPI ${dec2(f.spi)}).` });
  }
  for (const p of partidas) {
    const s = serie[p.cod + '|' + corte];
    const exceso = s.ac - s.ev;
    if (exceso > Math.max(1000, 0.03 * p.bac)) {
      const obs = [...new Set(planillas.filter(x => x.cod === p.cod && x.obs).map(x => x.obs))].sort();
      alertas.push({ nivel: 'serious', tipo: 'Coste', texto: `${p.desc}: certificado ${eur(exceso)} por encima de lo entregado/instalado. ` + (obs.length ? obs.join('; ') + '.' : '') });
    }
  }
  for (const p of partidas) {
    if ((p.fuente === 'REVIT' || p.fuente === 'ALBARAN') && cantMedida(p, corte) > p.cmed + 1e-6) {
      const verbo = p.fuente === 'ALBARAN' ? 'entregadas' : 'instaladas';
      alertas.push({ nivel: 'serious', tipo: 'Alcance', texto: `${p.desc}: ${verbo} ${num2(cantMedida(p, corte))} ${p.ud} frente a ${num2(p.cmed)} presupuestadas. El exceso no suma valor ganado: revisar alcance o modelo.` });
    }
  }
  for (const c of contingencia) {
    if (c.estado !== 'Propuesto') continue;
    const sinColchon = cont_ini <= 0;
    alertas.push({ nivel: sinColchon ? 'serious' : 'warning', tipo: 'Contingencia',
                   texto: `Pendiente de aprobar: ${c.concepto} (${eur(c.imp)}).` + (sinColchon ? ' No hay contingencia confirmada: si se aprueba, sube el coste final.' : '') });
  }
  alertas.sort((a, b) => ORDEN[a.nivel] - ORDEN[b.nivel]);

  const D = {
    proyecto, corte, n_sem, horizonte, fecha_corte: iso(masDias(inicio, 7 * corte - 1)), inicio: iso(inicio),
    fin_plan: iso(fin_plan), fin_prev: fin_prev ? iso(fin_prev) : null,
    BAC, PV, EV, AC, SPI, CPI, EAC, VAC: BAC - EAC, TCPI, tcpi_inalcanzable, terminado, ES, SPIt, retraso, dur_prev,
    ext_pend, cont_ini, cont_cons, cont_prop, cont_disp, curva, actividades: filasAct, alertas, extras, contingencia,
    umbral, umbral_act, PAGADO, pendiente_pago: AC - PAGADO, generado: new Date().toISOString().slice(0, 10), avisos,
  };

  // ---------------------------------------------------------------- tablas para Power BI
  const fact_semanal = [];
  for (const p of partidas) {
    const prev = { pv: 0, ev: 0, ac: 0 };
    for (const w of SEMANAS) {
      const s = serie[p.cod + '|' + w];
      const fila = { cod_partida: p.cod, semana: w, pv_semana: r2(s.pv - prev.pv), pv_acum: r2(s.pv) };
      for (const k of ['ev', 'ac']) {
        if (k in s) { fila[k + '_semana'] = r2(s[k] - prev[k]); fila[k + '_acum'] = r2(s[k]); prev[k] = s[k]; }
        else { fila[k + '_semana'] = null; fila[k + '_acum'] = null; }
      }
      prev.pv = s.pv;
      fact_semanal.push(fila);
    }
  }
  const tablas = {
    dim_partida: partidas.map(p => ({ cod_partida: p.cod, capitulo: p.cap, descripcion: p.desc, id_actividad: p.act, fuente_avance: p.fuente,
                                      ud_medicion: p.ud, cant_medicion: p.cmed, bac: r2(p.bac), perfil_pv: p.perfil, semana_entrega: p.perfil === 'ENTREGA' ? p.sem_ent : null })),
    dim_actividad: actividades.map(a => ({ id_actividad: a.id, actividad: a.nombre, sem_inicio: a.si, sem_fin: a.sf,
                                           fecha_inicio: iso(masDias(inicio, 7 * (a.si - 1))), fecha_fin: iso(finSem(a.sf)), predecesora: a.pred })),
    dim_semana: SEMANAS.map(w => ({ semana: w, fecha_inicio: iso(masDias(inicio, 7 * (w - 1))), fecha_fin: iso(masDias(inicio, 7 * w - 1)), es_pasada: w <= corte ? 1 : 0 })),
    fact_semanal,
    fact_planillas: planillas.map(x => ({ n_planilla: x.n, semana: x.sem, cod_partida: x.cod, cant_certificada: x.q, precio_aplicado: x.pu, importe: x.imp, observacion: x.obs, semana_pago: x.sem_pago })),
    extras: extras.map(e => ({ cod: e.cod, descripcion: e.desc, importe: e.imp, estado: e.estado, semana_limite: e.limite, motivo: e.motivo })),
    contingencia: contingencia.map(c => ({ semana: c.sem, concepto: c.concepto, importe: c.imp, estado: c.estado })),
    parametros: [{ semana_corte: corte, n_semanas: n_sem, fecha_inicio: iso(inicio), contingencia_pct: cont_pct, umbral, umbral_actividad: umbral_act }],
  };
  return { D, tablas, avisos };
}
