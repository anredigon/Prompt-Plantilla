/* Prueba de la app (requiere Node + Playwright):  node scripts/probar_app.js
 * Carga ejemplo/Ejemplo_Control_Proyecto.xlsx, compara con ejemplo/valores_esperados.json,
 * regenera ejemplo/Informe_Ejemplo_S11.html y ejemplo/PowerBI_Ejemplo.xlsx y prueba la plantilla vacía. */
const { chromium } = require('playwright');
const fs = require('fs'), path = require('path');
const base = path.resolve(__dirname, '..');
(async () => {
  const b = await chromium.launch(); const ctx = await b.newContext({ acceptDownloads: true });
  const p = await ctx.newPage(); const errores = []; p.on('pageerror', e => errores.push(e.message));
  await p.goto('file://' + path.join(base, 'App_Control_Proyecto.html'));
  await p.setInputFiles('#archivo', path.join(base, 'ejemplo/Ejemplo_Control_Proyecto.xlsx'));
  await p.waitForSelector('#app:not(.oculto) #curva svg');
  const D = await p.evaluate(() => resultado.D);
  const esperado = JSON.parse(fs.readFileSync(path.join(base, 'ejemplo/valores_esperados.json'), 'utf8'));
  let fallos = 0;
  for (const [k, v] of Object.entries(esperado)) {
    const ok = typeof v === 'number' ? Math.abs(D[k] - v) < 0.01 : D[k] === v;
    if (!ok) { fallos++; console.log('FALLO', k, 'esperado', v, 'obtenido', D[k]); }
  }
  const [inf] = await Promise.all([p.waitForEvent('download'), p.click('#exportar')]);
  await inf.saveAs(path.join(base, 'ejemplo/Informe_Ejemplo_S11.html'));
  const [pbi] = await Promise.all([p.waitForEvent('download'), p.click('#powerbi')]);
  await pbi.saveAs(path.join(base, 'ejemplo/PowerBI_Ejemplo.xlsx'));
  await p.click('#otro');
  await p.setInputFiles('#archivo', path.join(base, 'plantilla/Plantilla_Control_Proyecto.xlsx'));
  await p.waitForSelector('#error:not(.oculto)');
  const q = await ctx.newPage(); q.on('pageerror', e => errores.push('informe: ' + e.message));
  await q.goto('file://' + path.join(base, 'ejemplo/Informe_Ejemplo_S11.html'));
  await q.waitForSelector('#app:not(.oculto) #curva svg');
  if (errores.length) { fallos++; console.log('Errores JS:', errores); }
  console.log(fallos ? `${fallos} FALLOS` : `OK: ${Object.keys(esperado).length} valores coinciden; informe y Power BI regenerados.`);
  await b.close(); process.exit(fallos ? 1 : 0);
})();
