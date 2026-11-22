/* Informe del directorio: formatos, semáforos y render(D, modo). Compartido por la app sin conexión y la de claude.ai. */
const $ = (s, el = document) => el.querySelector(s);
const esc = v => String(v ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const signo = v => (v >= 0 ? '+' : '');
const eur = (v, d = 0) => { if (v == null) return '–'; const r = Math.round(v * 10 ** d) / 10 ** d; return new Intl.NumberFormat('es-ES', { style: 'currency', currency: 'EUR', maximumFractionDigits: d, minimumFractionDigits: d, useGrouping: 'always' }).format(r === 0 ? 0 : r); };
const keur = v => v == null ? '–' : (Math.abs(v) >= 1e6 ? (v / 1e6).toLocaleString('es-ES', { maximumFractionDigits: 2 }) + ' M€' : Math.round(v / 1e3).toLocaleString('es-ES') + ' k€');
const pct = (v, d = 0) => v == null || !isFinite(v) ? '–' : (v * 100).toLocaleString('es-ES', { maximumFractionDigits: d, minimumFractionDigits: d }) + ' %';
const dec = (v, d = 2) => v == null ? '–' : v.toLocaleString('es-ES', { maximumFractionDigits: d, minimumFractionDigits: d });
const fecha = iso => !iso ? '–' : new Date(iso + 'T12:00:00').toLocaleDateString('es-ES', { day: '2-digit', month: 'short', year: 'numeric' });
const STATUS = {
  good: { lab: 'OK', ic: '✓' }, warning: { lab: 'Vigilar', ic: '!' },
  serious: { lab: 'Desviado', ic: '▲' }, critical: { lab: 'Crítico', ic: '✕' }, sin_iniciar: { lab: 'No iniciada', ic: '–' }
};
const statusVar = s => s === 'sin_iniciar' ? 'var(--neutral)' : `var(--${s})`;
const badge = (s, lab) => `<span class="status"><span class="ic" style="background:${statusVar(s)};${s === 'warning' ? 'color:#0b0b0b' : ''}">${STATUS[s].ic}</span>${lab || STATUS[s].lab}</span>`;
const nivelPor = (v, u) => v == null ? 'sin_iniciar' : v >= u ? 'good' : v >= 0.85 ? 'warning' : v >= 0.7 ? 'serious' : 'critical';
const pasoBonito = bruto => { const p = 10 ** Math.floor(Math.log10(bruto || 1)); const m = bruto / p; return (m <= 1 ? 1 : m <= 2 ? 2 : m <= 2.5 ? 2.5 : m <= 5 ? 5 : 10) * p; };

const tt = $('#tt');
function showTT(html, ev) {
  tt.innerHTML = html; tt.style.opacity = 1;
  const w = tt.offsetWidth, h = tt.offsetHeight;
  let x = ev.clientX + 14, y = ev.clientY + 14;
  if (x + w > innerWidth - 8) x = ev.clientX - w - 14;
  if (y + h > innerHeight - 8) y = ev.clientY - h - 14;
  tt.style.left = x + 'px'; tt.style.top = y + 'px';
}
const hideTT = () => tt.style.opacity = 0;
const alerta = a => `<li><div>${badge(a.nivel)}<span class="tipo">${esc(a.tipo)}</span></div><div>${esc(a.texto)}</div></li>`;

// =================================================================== informe
function render(D, modo) {
  const nivelPlazo = nivelPor(D.SPIt, D.umbral);
  let nivelCoste = nivelPor(D.CPI, D.umbral);
  if (nivelCoste === 'good' && D.EAC > D.BAC) nivelCoste = 'warning'; // sobrecoste previsto nunca en verde
  const desvEAC = D.EAC - D.BAC;
  const enApp = modo === 'app';
  const conBarra = modo !== 'incrustado';  // incrustado: lo usa la app de claude.ai, que pone su propia barra
  const app = $('#app');
  app.innerHTML = `
<header>
  <div>
    <div class="sub">Informe semanal al directorio · ${esc(D.proyecto)}</div>
    <h1>Semana ${D.corte} de ${D.n_sem} · corte ${fecha(D.fecha_corte)}</h1>
  </div>
  ${conBarra ? `<div class="toolbar">
    ${enApp ? `<label>Semana de corte <input type="number" id="corte" min="1" max="104" value="${D.corte}"></label>
    <button class="theme" id="otro">Abrir otro proyecto</button>
    <button class="theme primario" id="exportar">Exportar informe</button>
    <button class="theme" id="powerbi">Datos Power BI</button>` : ''}
    <button class="theme" id="theme">Tema claro / oscuro</button><button class="theme" onclick="print()">Imprimir / PDF</button>
  </div>` : ''}
</header>
${D.avisos && D.avisos.length ? `<div class="avisos"><b>Revisa el Excel (${D.avisos.length} aviso${D.avisos.length > 1 ? 's' : ''}):</b><ul>${D.avisos.map(a => `<li>${esc(a)}</li>`).join('')}</ul></div>` : ''}

<section class="heads">
  <div class="card head">
    <div class="lab">${D.terminado ? 'Fin real' : 'Fin previsto al ritmo actual'}</div>
    <div class="big">${fecha(D.fin_prev)}</div>
    <div class="det">${badge(nivelPlazo)} &nbsp;${D.dur_prev == null ? 'Aún sin avance medible. Fin planificado ' + fecha(D.fin_plan) : signo(D.dur_prev - D.n_sem) + dec(D.dur_prev - D.n_sem, 1) + ' sem. sobre el plan (' + fecha(D.fin_plan) + ')'}. Retraso acumulado hoy: ${dec(D.retraso, 1)} sem.</div>
  </div>
  <div class="card head">
    <div class="lab">Coste previsto al final (EAC)</div>
    <div class="big">${eur(D.EAC)}</div>
    <div class="det">${badge(nivelCoste)} &nbsp;${signo(desvEAC)}${eur(desvEAC)} (${signo(desvEAC)}${pct(desvEAC / D.BAC, 1)}) sobre oferta</div>
  </div>
  <div class="card head">
    <div class="lab">Exposición máxima</div>
    <div class="big">${eur(D.EAC + D.ext_pend)}</div>
    <div class="det">EAC + ${eur(D.ext_pend)} de extras <b>sin decidir</b> fuera de la oferta</div>
  </div>
</section>

<section class="kpis">
  <div class="card kpi"><div class="lab">Avance real / planificado</div><div class="val">${pct(D.EV / D.BAC)} / ${pct(D.PV / D.BAC)}</div>
    <div class="meter" title="Planificado"><i style="width:${D.PV / D.BAC * 100}%;background:var(--plan-bar)"></i><i style="width:${D.EV / D.BAC * 100}%;height:4px;top:2px;background:var(--ev)"></i></div></div>
  <div class="card kpi"><div class="lab">SPI (€) · SPI(t)</div><div class="val">${dec(D.SPI)} · ${dec(D.SPIt)}</div><div class="det">${badge(nivelPor(D.SPI, D.umbral))}</div></div>
  <div class="card kpi"><div class="lab">CPI</div><div class="val">${dec(D.CPI)}</div><div class="det">${badge(nivelPor(D.CPI, D.umbral))}</div></div>
  <div class="card kpi"><div class="lab">Certificado − ganado</div><div class="val">${eur(D.AC - D.EV)}</div><div class="det">AC ${keur(D.AC)} · EV ${keur(D.EV)}<br>Pagado ${keur(D.PAGADO)} · pendiente ${keur(D.pendiente_pago)}</div></div>
  <div class="card kpi"><div class="lab">TCPI para cumplir oferta</div><div class="val">${D.tcpi_inalcanzable ? 'Inalcanzable' : dec(D.TCPI)}</div><div class="det">${D.tcpi_inalcanzable ? 'Ya se ha certificado toda la oferta y queda obra' : 'Eficiencia necesaria en lo que queda'}</div></div>
  <div class="card kpi"><div class="lab">Contingencia disponible</div><div class="val">${D.cont_ini > 0 ? eur(D.cont_disp) : 'Sin confirmar'}</div><div class="det">${D.cont_ini > 0 ? 'de ' + eur(D.cont_ini) + ' (' + pct(D.cont_disp / D.cont_ini) + ')' : 'No hay colchón: todo cambio aprobado sube el coste final'}</div></div>
</section>

<section class="grid2">
  <div class="card">
    <h2>Curva S · valor planificado, ganado y certificado (acumulado)</h2>
    <div class="legend">
      <span><i style="border-color:var(--pv);border-top-style:dashed"></i>Planificado (PV)</span>
      <span><i style="border-color:var(--ev)"></i>Ganado (EV)</span>
      <span><i style="border-color:var(--ac)"></i>Certificado (AC)</span>
    </div>
    <div id="curva"></div>
    <details><summary>Ver tabla de datos</summary><div id="curva-tabla"></div></details>
  </div>
  <div class="card">
    <h2>Decisiones y alertas para el directorio</h2>
    ${D.alertas.length ? `<ul class="alerts">${D.alertas.slice(0, 6).map(alerta).join('')}</ul>` : '<p class="muted">Sin alertas esta semana.</p>'}
    ${D.alertas.length > 6 ? `<details><summary>Ver ${D.alertas.length - 6} alertas más</summary><ul class="alerts" style="margin-top:10px">${D.alertas.slice(6).map(alerta).join('')}</ul></details>` : ''}
  </div>
</section>

<section class="card">
  <h2>Actividades · cronograma y avance a semana ${D.corte}</h2>
  <div class="legend">
    <span><i style="border-top:8px solid var(--plan-bar);width:16px"></i>Duración planificada</span>
    <span><i style="border-top:4px solid var(--ev);width:16px"></i>Avance real (% del presupuesto)</span>
    <span><i style="border-top:0;border-left:2px dashed var(--corte);height:12px;width:0"></i>Corte</span>
  </div>
  <div style="overflow-x:auto"><table id="tabla-act"></table></div>
</section>

<section class="cols2">
  <div class="card">
    <h2>Extras fuera de oferta</h2>
    ${D.extras.length ? `<table><thead><tr><th>Concepto</th><th class="n">Importe</th><th>Estado</th><th class="n">Límite</th></tr></thead>
    <tbody>${D.extras.map(e => `<tr><td>${esc(e.desc)}</td><td class="n">${eur(e.imp)}</td><td>${esc(e.estado)}</td><td class="n">${e.limite == null ? '–' : 'S' + e.limite}</td></tr>`).join('')}
    <tr><td><b>Total sin decidir</b></td><td class="n"><b>${eur(D.ext_pend)}</b></td><td colspan="2" class="muted">${pct(D.ext_pend / D.BAC, 1)} sobre la oferta</td></tr></tbody></table>` : '<p class="muted">No hay extras registrados.</p>'}
  </div>
  <div class="card">
    <h2>Contingencia</h2>
    <table><tbody>
      <tr><td>Inicial (${pct(D.cont_ini / D.BAC, 1)} s/ oferta${D.cont_ini > 0 ? '' : ': pendiente de confirmar por dirección'})</td><td class="n">${eur(D.cont_ini)}</td></tr>
      <tr><td>Consumida (aprobado)</td><td class="n">${eur(D.cont_cons)}</td></tr>
      <tr><td>Comprometida (propuesto)</td><td class="n">${eur(D.cont_prop)}</td></tr>
      <tr><td><b>Disponible</b></td><td class="n"><b>${eur(D.cont_disp)}</b></td></tr>
    </tbody></table>
    <p class="muted">${D.cont_ini > 0 ? `La contingencia no cubre los extras: ${eur(D.ext_pend)} pendientes frente a ${eur(D.cont_disp)} disponibles.` : `Sin contingencia confirmada, los ${eur(D.ext_pend)} de extras y cualquier cambio aprobado van directos al coste final.`}</p>
  </div>
</section>

<footer>
  Planificado: equipos al 100 % en su semana de entrega prevista; resto, lineal en la duración de su actividad.
  Valor ganado: equipos por albarán de entrega; resto de partidas del modelo, por unidades instaladas en Revit (CC_ESTADO); partidas fuera del modelo, % manual; montaje, por horas de los elementos instalados.
  Coste real: importes certificados en planillas (lo pagado se muestra aparte). Umbral de alerta: ${dec(D.umbral)} proyecto · ${dec(D.umbral_act)} actividades. SPI(t) por «earned schedule»: semana en que lo ganado hoy debía estar ganado (${dec(D.ES, 1)}) ÷ semana de corte.
  EAC = BAC / CPI. Generado ${fecha(D.generado)}.
</footer>`;

  // ---------------------------------------------------------------- curva S
  (function () {
    const host = $('#curva'), W = host.clientWidth || 700, H = 320, m = { l: 64, r: 92, t: 12, b: 34 };
    const iw = W - m.l - m.r, ih = H - m.t - m.b, n = D.horizonte || D.n_sem;
    const top = Math.max(D.BAC, ...D.curva.map(c => Math.max(c.ac || 0, c.pv))) || 1;
    const paso = pasoBonito(top / 6), maxY = Math.ceil(top / paso) * paso;
    const pasoX = Math.max(1, Math.ceil(n / 6));
    const x = w => m.l + (w / n) * iw, y = v => m.t + ih - (v / maxY) * ih;
    const pts = k => [[0, 0], ...D.curva.filter(c => c[k] != null).map(c => [c.sem, c[k]])];
    const path = k => pts(k).map((p, i) => (i ? 'L' : 'M') + x(p[0]).toFixed(1) + ',' + y(p[1]).toFixed(1)).join('');
    let g = '';
    for (let v = 0; v <= maxY + 1e-6; v += paso) g += `<line x1="${m.l}" x2="${m.l + iw}" y1="${y(v)}" y2="${y(v)}" stroke="var(--grid)"/><text x="${m.l - 8}" y="${y(v) + 4}" text-anchor="end" font-size="11" fill="var(--text-muted)">${keur(v)}</text>`;
    for (let w = 0; w <= n; w += pasoX) g += `<text x="${x(w)}" y="${H - 12}" text-anchor="middle" font-size="11" fill="var(--text-muted)">S${w}</text>`;
    const last = D.curva[D.corte - 1], fin = D.curva[n - 1];
    const lbl = (w, v, txt, dy = 0) => `<text x="${x(w) + 8}" y="${y(v) + 4 + dy}" font-size="11" fill="var(--text-secondary)">${txt}</text>`;
    const dyEV = last.ev == null || last.ac == null ? 0 : Math.abs(y(last.ev) - y(last.ac)) < 14 ? (last.ev < last.ac ? 8 : -8) : 0;
    host.innerHTML = `<svg width="${W}" height="${H}" role="img" aria-label="Curva S de PV, EV y AC">
      ${g}
      <line x1="${m.l}" x2="${m.l + iw}" y1="${y(D.BAC)}" y2="${y(D.BAC)}" stroke="var(--text-muted)" stroke-dasharray="2 3"/>
      <text x="${m.l + 4}" y="${y(D.BAC) - 5}" font-size="11" fill="var(--text-muted)">Oferta ${keur(D.BAC)}</text>
      <line x1="${x(D.corte)}" x2="${x(D.corte)}" y1="${m.t}" y2="${m.t + ih}" stroke="var(--corte)" stroke-dasharray="4 3"/>
      <text x="${x(D.corte) + 4}" y="${m.t + 10}" font-size="11" fill="var(--text-secondary)">Corte S${D.corte}</text>
      <path d="${path('pv')}" fill="none" stroke="var(--pv)" stroke-width="2" stroke-dasharray="6 4" stroke-linejoin="round"/>
      <path d="${path('ac')}" fill="none" stroke="var(--ac)" stroke-width="2" stroke-linejoin="round"/>
      <path d="${path('ev')}" fill="none" stroke="var(--ev)" stroke-width="2" stroke-linejoin="round"/>
      ${lbl(n, fin.pv, 'PV')}
      ${last.ac != null ? lbl(D.corte, last.ac, 'AC ' + keur(last.ac), -dyEV) : ''}
      ${last.ev != null ? lbl(D.corte, last.ev, 'EV ' + keur(last.ev), dyEV) : ''}
      <g id="xh" style="display:none"><line id="xhl" y1="${m.t}" y2="${m.t + ih}" stroke="var(--text-muted)"/>
        <circle id="c-pv" r="4" fill="var(--pv)" stroke="var(--surface-1)" stroke-width="2"/>
        <circle id="c-ev" r="4" fill="var(--ev)" stroke="var(--surface-1)" stroke-width="2"/>
        <circle id="c-ac" r="4" fill="var(--ac)" stroke="var(--surface-1)" stroke-width="2"/></g>
      <rect id="hit" x="${m.l}" y="${m.t}" width="${iw}" height="${ih}" fill="transparent"/>
    </svg>`;
    const svg = host.querySelector('svg');
    $('#hit', svg).addEventListener('mousemove', ev => {
      const r = svg.getBoundingClientRect();
      const w = Math.min(n, Math.max(1, Math.round((ev.clientX - r.left - m.l) / iw * n)));
      const c = D.curva[w - 1];
      $('#xh', svg).style.display = '';
      $('#xhl', svg).setAttribute('x1', x(w)); $('#xhl', svg).setAttribute('x2', x(w));
      for (const k of ['pv', 'ev', 'ac']) {
        const el = $('#c-' + k, svg);
        if (c[k] == null) { el.style.display = 'none'; continue; }
        el.style.display = ''; el.setAttribute('cx', x(w)); el.setAttribute('cy', y(c[k]));
      }
      const row = (k, lab) => c[k] == null ? '' : `<div class="r"><span><span class="sw" style="background:var(--${k})"></span>${lab}</span><span>${eur(c[k])}</span></div>`;
      showTT(`<b>Semana ${w}</b>${row('pv', 'Planificado')}${row('ev', 'Ganado')}${row('ac', 'Certificado')}`, ev);
    });
    $('#hit', svg).addEventListener('mouseleave', () => { $('#xh', svg).style.display = 'none'; hideTT(); });
    $('#curva-tabla').innerHTML = `<table><thead><tr><th>Semana</th><th class="n">PV</th><th class="n">EV</th><th class="n">AC</th></tr></thead><tbody>${D.curva.map(c => `<tr><td>S${c.sem}</td><td class="n">${eur(c.pv)}</td><td class="n">${eur(c.ev)}</td><td class="n">${eur(c.ac)}</td></tr>`).join('')}</tbody></table>`;
  })();

  // ---------------------------------------------------------------- actividades con gantt
  (function () {
    const n = D.horizonte || D.n_sem, GW = 300, GH = 22;
    const gx = w => (w / n) * GW;
    const filas = D.actividades.map((a, i) => {
      const x0 = gx(a.si - 1), x1 = gx(a.sf), wReal = (x1 - x0) * Math.min(1, a.pct_real);
      const gantt = `<svg width="${GW}" height="${GH}" style="display:block">
        <rect x="${x0}" y="4" width="${Math.max(2, x1 - x0 - 2)}" height="14" rx="4" fill="var(--plan-bar)"/>
        ${wReal > 0 ? `<rect x="${x0}" y="9" width="${Math.max(2, wReal - 2)}" height="4" rx="2" fill="var(--ev)"/>` : ''}
        <line x1="${gx(D.corte)}" x2="${gx(D.corte)}" y1="0" y2="${GH}" stroke="var(--corte)" stroke-dasharray="3 2"/>
      </svg>`;
      return `<tr data-i="${i}">
        <td><b>${esc(a.id)}</b><br><span class="muted">${esc(a.nombre)}</span></td>
        <td class="gantt-cell">${gantt}<span class="muted">S${a.si}–S${a.sf}</span></td>
        <td class="n">${pct(a.pct_plan)}</td><td class="n">${pct(a.pct_real)}</td>
        <td class="n">${dec(a.spi)}</td><td class="n">${dec(a.cpi)}</td>
        <td class="n">${eur(a.ev - a.ac)}</td><td>${badge(a.estado)}</td></tr>`;
    }).join('');
    const t = $('#tabla-act');
    t.innerHTML = `<thead><tr><th>Actividad</th><th>S1 → S${n}</th><th class="n">% plan.</th><th class="n">% real</th><th class="n">SPI</th><th class="n">CPI</th><th class="n">Ganado − certif.</th><th>Estado</th></tr></thead><tbody>${filas}</tbody>`;
    t.querySelectorAll('tbody tr').forEach(tr => {
      const a = D.actividades[+tr.dataset.i];
      tr.addEventListener('mousemove', ev => showTT(`<b>${esc(a.id)} · ${esc(a.nombre)}</b>
        <div class="r"><span>Presupuesto</span><span>${eur(a.bac)}</span></div>
        <div class="r"><span><span class="sw" style="background:var(--pv)"></span>Planificado</span><span>${eur(a.pv)}</span></div>
        <div class="r"><span><span class="sw" style="background:var(--ev)"></span>Ganado</span><span>${eur(a.ev)}</span></div>
        <div class="r"><span><span class="sw" style="background:var(--ac)"></span>Certificado</span><span>${eur(a.ac)}</span></div>`, ev));
      tr.addEventListener('mouseleave', hideTT);
    });
  })();
}
