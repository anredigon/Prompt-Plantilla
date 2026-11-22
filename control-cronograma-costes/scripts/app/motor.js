/* Motor de cálculo de valor ganado (EVM) a partir del Excel de la plantilla.
 * calcular(wb, {corte}) -> { D, tablas, avisos }
 *   wb: libro leído con SheetJS (XLSX.read(buffer, {type: 'array'})).
 *   D: datos del informe; tablas: tablas para Power BI; avisos: problemas de datos encontrados.
 * Reglas (validadas por dirección): ver README, sección "Reglas de cálculo".
 */
function calcular(wb, opciones) {
  opciones = opciones || {};
  const avisos = [];
  const HOJAS = ['Parametros', 'Cronograma', 'Oferta', 'Revit_Export', 'Suministros', 'Horas_Montaje',
                 'Avance_Manual', 'Planillas', 'Extras', 'Contingencia'];
  const faltan = HOJAS.filter(h => !wb.Sheets[h]);
  if (faltan.length) throw new Error('El Excel no tiene la estructura de la plantilla. Faltan las hojas: ' + faltan.join(', '));

  const filasDe = h => XLSX.utils.sheet_to_json(wb.Sheets[h], { header: 1, raw: true, defval: null }).slice(1);
  const valor = (h, ref) => { const c = wb.Sheets[h][ref]; return c ? c.v : null; };
  const num = v => {
    if (typeof v === 'number' && isFinite(v)) return v;
    if (typeof v === 'string' && v.trim() !== '') { const n = Number(v.trim().replace(/\./g, '').replace(',', '.')); return isFinite(n) ? n : null; }
    return null;
  };
  const txt = v => (v == null ? '' : String(v).trim());
  const r2 = x => Math.round(x * 100) / 100;
  const div = (a, b) => (b ? a / b : null);
  const MS_DIA = 864e5;
  const serialAFecha = v => {
    if (typeof v === 'number') return new Date(Math.round((v - 25569) * MS_DIA));
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
  let corte = Math.round(num(opciones.corte != null ? opciones.corte : valor('Parametros', 'B6')) || 1);
  corte = Math.min(Math.max(corte, 1), n_sem);
  const cont_pct = num(valor('Parametros', 'B8')) || 0;
  const umbral = num(valor('Parametros', 'B9')) || 0.95;
  const umbral_act = num(valor('Parametros', 'B10')) || 0.9;

  // ---------------------------------------------------------------- lectura
  const actividades = [];
  for (const r of filasDe('Cronograma')) {
    if (!txt(r[0])) continue;
    const si = num(r[2]), sf = num(r[3]);
    if (si == null || sf == null || sf < si) { avisos.push(`Cronograma: la actividad ${txt(r[0])} no tiene semanas de inicio/fin válidas; se ignora.`); continue; }
    actividades.push({ id: txt(r[0]), nombre: txt(r[1]), si: Math.round(si), sf: Math.round(sf), pred: txt(r[7]) });
  }
  const actPorId = Object.fromEntries(actividades.map(a => [a.id, a]));

  const FUENTES = ['ALBARAN', 'REVIT', 'MANUAL', 'DERIVADO'];
  const partidas = [];
  for (const r of filasDe('Oferta')) {
    const cod = txt(r[0]);
    if (!cod || num(r[4]) == null || !txt(r[6])) continue;
    const a = actPorId[txt(r[6])];
    if (!a) { avisos.push(`Oferta: la partida ${cod} apunta a la actividad "${txt(r[6])}", que no está en el Cronograma; se ignora.`); continue; }
    const fuente = txt(r[7]).toUpperCase();
    if (!FUENTES.includes(fuente)) { avisos.push(`Oferta: la partida ${cod} tiene una fuente de avance no válida ("${txt(r[7])}"); se ignora.`); continue; }
    const cmed = num(r[9]);
    if ((fuente === 'REVIT' || fuente === 'ALBARAN') && !(cmed > 0)) { avisos.push(`Oferta: la partida ${cod} necesita CANT. MEDICIÓN mayor que 0; se ignora.`); continue; }
    const perfil = txt(r[22]).toUpperCase() === 'ENTREGA' ? 'ENTREGA' : 'LINEAL';
    const sem_ent = num(r[23]) != null ? Math.round(num(r[23])) : a.si;
    partidas.push({ cod, cap: txt(r[1]), desc: txt(r[2]), bac: num(r[4]), act: a.id, fuente, ud: txt(r[8]),
                    cmed: cmed || 1, perfil, sem_ent });
  }
  if (!partidas.length) throw new Error('La hoja Oferta no tiene ninguna partida válida (código, importe y actividad).');
  const codigos = new Set(partidas.map(p => p.cod));

  const revit = [];
  for (const r of filasDe('Revit_Export')) {
    const cod = txt(r[4]);
    if (!cod) continue;
    const sem = num(r[9]);
    revit.push({ cod, q: num(r[6]) || 0, estado: txt(r[8]), sem: sem == null ? null : Math.round(sem) });
  }
  const huerfanos = new Set(revit.filter(e => !codigos.has(e.cod)).map(e => e.cod));
  if (huerfanos.size) avisos.push(`Revit_Export: hay elementos con CC_COD_PARTIDA que no existe en la Oferta: ${[...huerfanos].slice(0, 8).join(', ')}.`);

  const manual = {};
  for (const r of filasDe('Avance_Manual')) {
    if (!txt(r[0])) continue;
    manual[txt(r[0])] = r.slice(2, 2 + n_sem).map(v => num(v));
  }
  const suministros = [];
  for (const r of filasDe('Suministros')) {
    if (num(r[0]) == null || !txt(r[1])) continue;
    suministros.push({ sem: Math.round(num(r[0])), cod: txt(r[1]), q: num(r[2]) || 0 });
  }
  const horas = {};
  for (const r of filasDe('Horas_Montaje')) if (txt(r[0]) && num(r[2]) != null) horas[txt(r[0])] = num(r[2]);

  const planillas = [];
  for (const r of filasDe('Planillas')) {
    if (num(r[1]) == null || !txt(r[2]) || num(r[3]) == null || num(r[4]) == null) continue;
    const pago = num(r[7]);
    planillas.push({ n: txt(r[0]), sem: Math.round(num(r[1])), cod: txt(r[2]), q: num(r[3]), pu: num(r[4]),
                     imp: r2(num(r[3]) * num(r[4])), obs: txt(r[6]), sem_pago: pago == null ? null : Math.round(pago) });
  }
  const sinPartida = new Set(planillas.filter(x => !codigos.has(x.cod)).map(x => x.cod));
  if (sinPartida.size) avisos.push(`Planillas: hay líneas con partidas que no existen en la Oferta (no cuentan en el coste): ${[...sinPartida].join(', ')}.`);

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

  // ---------------------------------------------------------------- cálculo semanal
  const SEMANAS = Array.from({ length: n_sem }, (_, i) => i + 1);
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
  if (partidas.some(p => p.fuente === 'DERIVADO') && !horasTot) avisos.push('Hay partidas DERIVADO (montaje) pero la hoja Horas_Montaje no tiene horas para elementos del modelo: su avance será 0.');
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

  // Earned schedule (Lipke): C = última semana con PV acumulado <= EV; se interpola con la siguiente
  const pvs = [0, ...curva.map(c => c.pv)];
  let C = 0;
  for (let t = 0; t < pvs.length; t++) if (pvs[t] <= EV + 1e-6) C = t;
  const ES = C === pvs.length - 1 ? C : C + (EV - pvs[C]) / (pvs[C + 1] - pvs[C]);
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
  const PAGADO = planillas.filter(x => codigos.has(x.cod) && x.sem_pago != null && x.sem_pago <= corte).reduce((s, x) => s + x.imp, 0);

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
    proyecto, corte, n_sem, fecha_corte: iso(masDias(inicio, 7 * corte - 1)), inicio: iso(inicio),
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
