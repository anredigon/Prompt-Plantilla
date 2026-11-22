"""
Lee Control_Proyecto_Maqueta.xlsx (solo celdas de entrada) y genera:
  - dashboard_directorio.html  (informe semanal autocontenido)
  - powerbi/*.csv              (modelo en estrella para Power BI)

Uso:  python generar_dashboard.py [ruta_excel] [carpeta_salida]
"""
import csv
import json
import sys
from datetime import date, datetime, timedelta
from pathlib import Path

from openpyxl import load_workbook

BASE = Path(__file__).resolve().parent.parent
XLSX = Path(sys.argv[1]) if len(sys.argv) > 1 else BASE / "datos" / "Control_Proyecto_Maqueta.xlsx"
OUT = Path(sys.argv[2]) if len(sys.argv) > 2 else BASE


def filas(ws, desde=2):
    for row in ws.iter_rows(min_row=desde, values_only=True):
        if row and row[0] not in (None, ""):
            yield row


def as_date(v):
    return v.date() if isinstance(v, datetime) else v


# ------------------------------------------------------------------ lectura
wb = load_workbook(XLSX, data_only=False)
P = wb["Parametros"]
proyecto = P["B3"].value
inicio = as_date(P["B4"].value)
n_sem = int(P["B5"].value)
corte = int(P["B6"].value)
cont_pct = float(P["B8"].value)
umbral = float(P["B9"].value)

actividades = []
for r in filas(wb["Cronograma"]):
    actividades.append(dict(id=r[0], nombre=r[1], si=int(r[2]), sf=int(r[3]), pred=r[7] or ""))
act_by_id = {a["id"]: a for a in actividades}

partidas = []
for r in filas(wb["Oferta"]):
    # Cualquier partida con código, importe numérico y actividad (la fila TOTAL no tiene código)
    if not isinstance(r[4], (int, float)) or not r[6]:
        continue
    partidas.append(dict(cod=r[0], cap=r[1], desc=r[2], cant=r[3], bac=float(r[4]), act=r[6],
                         fuente=r[7], ud=r[8], cmed=float(r[9])))

revit = []
for r in filas(wb["Revit_Export"]):
    revit.append(dict(eid=r[0], cat=r[1], tipo=r[2], zona=r[3], cod=r[4], act=r[5], q=float(r[6] or 0),
                      ud=r[7], estado=r[8], sem=int(r[9]) if r[9] not in (None, "") else None))

manual = {}
for r in filas(wb["Avance_Manual"]):
    if r[0] == "Introducir el % ACUMULADO al final de cada semana (0 % a 100 %). Columnas futuras en blanco.":
        continue
    manual[r[0]] = [float(v) if isinstance(v, (int, float)) else None for v in r[2:2 + n_sem]]

planillas = []
for r in filas(wb["Planillas"]):
    planillas.append(dict(n=r[0], sem=int(r[1]), cod=r[2], q=float(r[3]), pu=float(r[4]),
                          imp=round(float(r[3]) * float(r[4]), 2), obs=r[6] or ""))

extras = []
for r in filas(wb["Extras"]):
    if str(r[0]).startswith("X"):
        extras.append(dict(cod=r[0], desc=r[1], imp=float(r[2]), estado=r[3], limite=int(r[4]), motivo=r[5]))

contingencia = []
for r in filas(wb["Contingencia"]):
    if isinstance(r[0], (int, float)):
        contingencia.append(dict(sem=int(r[0]), concepto=r[1], imp=float(r[2]), estado=r[3]))

# ------------------------------------------------------------------ cálculo semanal
SEMANAS = list(range(1, n_sem + 1))


def pv_acum(p, w):
    a = act_by_id[p["act"]]
    return p["bac"] * min(1.0, max(0.0, (w - a["si"] + 1) / (a["sf"] - a["si"] + 1)))


def pct_revit(p, w):
    q = sum(e["q"] for e in revit if e["cod"] == p["cod"] and e["estado"] == "Instalado"
            and e["sem"] is not None and e["sem"] <= w)
    return q / p["cmed"]


def frac_elementos(cap, w):
    cods = {p["cod"] for p in partidas if p["cap"] == cap and p["fuente"] == "REVIT"}
    tot = sum(1 for e in revit if e["cod"] in cods)
    hechos = sum(1 for e in revit if e["cod"] in cods and e["estado"] == "Instalado"
                 and e["sem"] is not None and e["sem"] <= w)
    return hechos / tot if tot else 0.0


def pct_avance(p, w):
    if p["fuente"] == "REVIT":
        return pct_revit(p, w)
    if p["fuente"] == "MANUAL":
        serie = manual.get(p["cod"], [])
        v = serie[w - 1] if w - 1 < len(serie) else None
        return v or 0.0
    return frac_elementos(p["cap"], w)


serie = {}  # (cod, w) -> dict
for p in partidas:
    for w in SEMANAS:
        d = dict(pv=pv_acum(p, w))
        if w <= corte:
            d["ev"] = p["bac"] * pct_avance(p, w)
            d["ac"] = sum(x["imp"] for x in planillas if x["cod"] == p["cod"] and x["sem"] <= w)
        serie[(p["cod"], w)] = d


def tot(campo, w, cods=None):
    return sum(serie[(p["cod"], w)].get(campo, 0.0) for p in partidas if cods is None or p["cod"] in cods)


BAC = sum(p["bac"] for p in partidas)
PV, EV, AC = tot("pv", corte), tot("ev", corte), tot("ac", corte)
def div(a, b):
    return a / b if b else None


SPI, CPI = div(EV, PV), div(EV, AC)
# Sin rendimiento de coste medible (CPI nulo o 0): EAC = AC + trabajo pendiente a precio de oferta
EAC = BAC / CPI if CPI else AC + (BAC - EV)
curva = [dict(sem=w, pv=round(tot("pv", w), 2),
              ev=round(tot("ev", w), 2) if w <= corte else None,
              ac=round(tot("ac", w), 2) if w <= corte else None) for w in SEMANAS]

# Earned schedule: momento en que lo ganado hoy debía estar ganado
pvs = [0.0] + [c["pv"] for c in curva]
ES = float(len(pvs) - 1)  # si lo ganado supera todo el PV, se ha ganado el plan completo
for t in range(1, len(pvs)):
    if pvs[t] >= EV:
        ES = (t - 1) + (EV - pvs[t - 1]) / (pvs[t] - pvs[t - 1]) if pvs[t] > pvs[t - 1] else t - 1
        break
SPIt = div(ES, corte) if PV > 0 else None  # sin nada planificado aún no hay índice de plazo
dur_prev = div(n_sem, SPIt)
retraso_sem = corte - ES


def fecha_fin_sem(w):
    return inicio + timedelta(days=7 * w - 3)


fin_plan = fecha_fin_sem(n_sem)
fin_prev = inicio + timedelta(days=round(7 * dur_prev) - 3) if dur_prev else None

ext_pend = sum(e["imp"] for e in extras if e["estado"] == "Sin decidir")
cont_ini = cont_pct * BAC
cont_cons = sum(c["imp"] for c in contingencia if c["estado"] == "Aprobado")
cont_prop = sum(c["imp"] for c in contingencia if c["estado"] == "Propuesto")
cont_disp = cont_ini - cont_cons - cont_prop

# ------------------------------------------------------------------ por actividad
ORDEN = {"critical": 0, "serious": 1, "warning": 2, "good": 3}


def nivel(indice):
    """Mismo umbral en tabla, alertas y titulares: >=umbral OK, >=0,85 vigilar, >=0,70 desviado, <0,70 crítico."""
    if indice is None or indice >= umbral:
        return "good"
    return "warning" if indice >= 0.85 else ("serious" if indice >= 0.7 else "critical")


filas_act = []
for a in actividades:
    cods = {p["cod"] for p in partidas if p["act"] == a["id"]}
    bac = sum(p["bac"] for p in partidas if p["cod"] in cods)
    pv, ev, ac = tot("pv", corte, cods), tot("ev", corte, cods), tot("ac", corte, cods)
    spi = ev / pv if pv else None
    cpi = ev / ac if ac else None
    if pv == 0 and ev == 0 and ac == 0:
        estado = "sin_iniciar"
    else:
        estado = min(nivel(spi), nivel(cpi), key=lambda k: ORDEN[k])
    filas_act.append(dict(id=a["id"], nombre=a["nombre"], si=a["si"], sf=a["sf"], bac=round(bac, 2),
                          pv=round(pv, 2), ev=round(ev, 2), ac=round(ac, 2),
                          pct_plan=pv / bac if bac else 0, pct_real=ev / bac if bac else 0,
                          spi=spi, cpi=cpi, estado=estado))

# ------------------------------------------------------------------ alertas
def eur(x):
    return f"{x:,.0f} €".replace(",", ".")


def dec(x):
    return f"{x:.2f}".replace(".", ",")


alertas = []
for e in extras:
    if e["estado"] == "Sin decidir":
        margen = e["limite"] - corte
        alertas.append(dict(nivel="critical" if margen <= 2 else ("serious" if margen <= 6 else "warning"),
                            tipo="Decisión",
                            texto=f"{e['desc']} ({eur(e['imp'])}) sin decidir. Límite: semana {e['limite']} "
                                  f"(quedan {margen} sem.). {e['motivo']}"))
for f in filas_act:
    if f["pv"] > 0 and f["pct_plan"] >= 0.999 and f["pct_real"] < 0.05:
        alertas.append(dict(nivel="critical", tipo="Plazo",
                            texto=f"{f['nombre']}: debía estar terminada en la semana {f['sf']} y el avance es "
                                  f"{f['pct_real']:.0%}."))
    elif f["spi"] is not None and nivel(f["spi"]) in ("serious", "critical"):
        alertas.append(dict(nivel=nivel(f["spi"]), tipo="Plazo",
                            texto=f"{f['nombre']}: avance {f['pct_real']:.0%} frente a {f['pct_plan']:.0%} "
                                  f"planificado (SPI {dec(f['spi'])})."))
for p in partidas:
    s = serie[(p["cod"], corte)]
    exceso = s["ac"] - s["ev"]
    if exceso > max(1000, 0.03 * p["bac"]):
        obs = sorted({x["obs"] for x in planillas if x["cod"] == p["cod"] and x["obs"]})
        alertas.append(dict(nivel="serious", tipo="Coste",
                            texto=f"{p['desc']}: certificado {eur(exceso)} por encima de lo instalado. "
                                  + ("; ".join(obs) + "." if obs else "")))
for c in contingencia:
    if c["estado"] == "Propuesto":
        alertas.append(dict(nivel="warning", tipo="Contingencia",
                            texto=f"Pendiente de aprobar contra contingencia: {c['concepto']} ({eur(c['imp'])})."))
alertas.sort(key=lambda x: ORDEN[x["nivel"]])



# ------------------------------------------------------------------ salida HTML
datos = dict(
    proyecto=proyecto, corte=corte, n_sem=n_sem,
    fecha_corte=(inicio + timedelta(days=7 * corte - 1)).isoformat(),
    inicio=inicio.isoformat(), fin_plan=fin_plan.isoformat(), fin_prev=fin_prev.isoformat() if fin_prev else None,
    BAC=BAC, PV=PV, EV=EV, AC=AC, SPI=SPI, CPI=CPI, EAC=EAC, VAC=BAC - EAC,
    TCPI=div(BAC - EV, BAC - AC), ES=ES, SPIt=SPIt, retraso=retraso_sem, dur_prev=dur_prev,
    ext_pend=ext_pend, cont_ini=cont_ini, cont_cons=cont_cons, cont_prop=cont_prop, cont_disp=cont_disp,
    curva=curva, actividades=filas_act, alertas=alertas, extras=extras, contingencia=contingencia,
    umbral=umbral, generado=date.today().isoformat(),
)
plantilla = (Path(__file__).resolve().parent / "plantilla_dashboard.html").read_text(encoding="utf-8")
js = json.dumps(datos, ensure_ascii=False, default=str).replace("</", "<\\/")  # evita cerrar <script>
html = plantilla.replace("/*__DATOS__*/null", js)
OUT.mkdir(parents=True, exist_ok=True)
(OUT / "dashboard_directorio.html").write_text(html, encoding="utf-8")

# ------------------------------------------------------------------ salida Power BI
pb = OUT / "powerbi"
pb.mkdir(parents=True, exist_ok=True)


def escribir(nombre, cab, rows):
    with open(pb / nombre, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh)
        w.writerow(cab)
        w.writerows(rows)


escribir("dim_partida.csv", ["cod_partida", "capitulo", "descripcion", "id_actividad", "fuente_avance",
                             "ud_medicion", "cant_medicion", "bac"],
         [[p["cod"], p["cap"], p["desc"], p["act"], p["fuente"], p["ud"], p["cmed"], round(p["bac"], 2)]
          for p in partidas])
escribir("dim_actividad.csv", ["id_actividad", "actividad", "sem_inicio", "sem_fin", "fecha_inicio",
                               "fecha_fin", "predecesora"],
         [[a["id"], a["nombre"], a["si"], a["sf"], (inicio + timedelta(days=7 * (a["si"] - 1))).isoformat(),
           fecha_fin_sem(a["sf"]).isoformat(), a["pred"]] for a in actividades])
escribir("dim_semana.csv", ["semana", "fecha_inicio", "fecha_fin", "es_pasada"],
         [[w, (inicio + timedelta(days=7 * (w - 1))).isoformat(), (inicio + timedelta(days=7 * w - 1)).isoformat(),
           int(w <= corte)] for w in SEMANAS])
rows = []
for p in partidas:
    prev = dict(pv=0.0, ev=0.0, ac=0.0)
    for w in SEMANAS:
        s = serie[(p["cod"], w)]
        r = [p["cod"], w, round(s["pv"] - prev["pv"], 2), round(s["pv"], 2)]
        for k in ("ev", "ac"):
            if k in s:
                r += [round(s[k] - prev[k], 2), round(s[k], 2)]
                prev[k] = s[k]
            else:
                r += ["", ""]
        prev["pv"] = s["pv"]
        rows.append(r)
escribir("fact_semanal.csv", ["cod_partida", "semana", "pv_semana", "pv_acum", "ev_semana", "ev_acum",
                              "ac_semana", "ac_acum"], rows)
escribir("fact_planillas.csv", ["n_planilla", "semana", "cod_partida", "cant_certificada", "precio_aplicado",
                                "importe", "observacion"],
         [[x["n"], x["sem"], x["cod"], x["q"], x["pu"], x["imp"], x["obs"]] for x in planillas])
escribir("extras.csv", ["cod", "descripcion", "importe", "estado", "semana_limite", "motivo"],
         [[e["cod"], e["desc"], e["imp"], e["estado"], e["limite"], e["motivo"]] for e in extras])
escribir("contingencia.csv", ["semana", "concepto", "importe", "estado"],
         [[c["sem"], c["concepto"], c["imp"], c["estado"]] for c in contingencia])
escribir("parametros.csv", ["semana_corte", "n_semanas", "fecha_inicio", "contingencia_pct", "umbral"],
         [[corte, n_sem, inicio.isoformat(), cont_pct, umbral]])

print(json.dumps(dict(BAC=round(BAC, 2), PV=round(PV, 2), EV=round(EV, 2), AC=round(AC, 2),
                      SPI=SPI and round(SPI, 4), CPI=CPI and round(CPI, 4), EAC=round(EAC, 2), ES=round(ES, 2),
                      retraso_sem=round(retraso_sem, 2), fin_prev=fin_prev and fin_prev.isoformat(),
                      alertas=len(alertas)), ensure_ascii=False, indent=1))
