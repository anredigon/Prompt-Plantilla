"""
Genera la PLANTILLA Excel de control semanal de cronograma y costes (valor ganado).

    python crear_plantilla.py                  -> plantilla/Plantilla_Control_Proyecto.xlsx  (en blanco)
    python crear_plantilla.py --ejemplo        -> ejemplo/Ejemplo_Control_Proyecto.xlsx      (maqueta rellena)
    python crear_plantilla.py --ejemplo salida.xlsx

La plantilla trae todas las fórmulas preparadas para proyectos de hasta 300 partidas, 150 actividades,
5.000 elementos de Revit, 2.000 líneas de planilla y 104 semanas. Solo se rellenan las celdas en azul.
"""
import argparse
from pathlib import Path

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

BASE = Path(__file__).resolve().parent.parent
ap = argparse.ArgumentParser()
ap.add_argument("salida", nargs="?")
ap.add_argument("--ejemplo", action="store_true", help="rellena con los datos de la maqueta")
args = ap.parse_args()
E = None
if args.ejemplo:
    import datos_ejemplo as E  # noqa: N812
SALIDA = Path(args.salida) if args.salida else (
    BASE / "ejemplo" / "Control_Ejemplo.xlsx" if E else BASE / "Plantilla_Control_Proyecto.xlsx")

# Capacidad de la plantilla
MAX_PART, MAX_ACT, MAX_REVIT, MAX_PLAN = 300, 150, 5000, 2000
MAX_SUM, MAX_HORAS, MAX_MAN, MAX_SEM, MAX_EXT, MAX_CONT = 1000, 300, 100, 104, 50, 100

# ---------------------------------------------------------------- estilos
F = "Arial"
f_base = Font(name=F, size=10)
f_input = Font(name=F, size=10, color="0000FF")
f_link = Font(name=F, size=10, color="008000")
f_head = Font(name=F, size=10, bold=True, color="FFFFFF")
f_title = Font(name=F, size=14, bold=True)
f_bold = Font(name=F, size=10, bold=True)
fill_head = PatternFill("solid", fgColor="1F3A5F")
fill_key = PatternFill("solid", fgColor="FFFF00")
fill_total = PatternFill("solid", fgColor="E7ECF2")
EUR = '#,##0.00 €;(#,##0.00 €);-'
EUR0 = '#,##0 €;(#,##0 €);-'
PCT = '0.0%;(0.0%);-'
RATIO = '0.00'
CORTE = "Parametros!$B$6"


def cabecera(ws, titulos, anchos):
    for c, t in enumerate(titulos, 1):
        cell = ws.cell(row=1, column=c, value=t)
        cell.font, cell.fill = f_head, fill_head
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 30
    for c, a in enumerate(anchos, 1):
        ws.column_dimensions[get_column_letter(c)].width = a
    ws.freeze_panes = "A2"


def put(ws, ref, valor, font=f_base, fmt=None):
    cell = ws[ref]
    cell.value = valor
    cell.font = font
    if fmt:
        cell.number_format = fmt
    return cell


def lista(ws, rango, valores):
    dv = DataValidation(type="list", formula1='"' + ",".join(valores) + '"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(rango)


wb = Workbook()

# ================================================================ LEEME
ws = wb.active
ws.title = "LEEME"
ws.column_dimensions["A"].width = 26
ws.column_dimensions["B"].width = 115
put(ws, "A1", "Control semanal de cronograma y costes (valor ganado) · plantilla", f_title)
filas = [
    ("Estado", "EJEMPLO con datos ficticios (costes = oferta 'Costes GPT 09 0008')." if E else
     "PLANTILLA EN BLANCO. Haz una copia por proyecto."),
    ("Texto azul", "Dato de entrada: se escribe a mano o se pega de una exportación."),
    ("Texto negro", "Fórmula. No sobrescribir."),
    ("Texto verde", "Fórmula que trae datos de otra hoja."),
    ("Relleno amarillo", "Supuesto clave a confirmar."),
    ("", ""),
    ("Proyecto nuevo", "1) Copia esta plantilla en la carpeta de Drive del proyecto y renómbrala (p. ej. Control_<Proyecto>.xlsx)."),
    ("", "2) 'Parametros': nombre, lunes de la semana 1, nº de semanas, contingencia y umbrales."),
    ("", "3) 'Cronograma': una fila por actividad (ID, nombre, semana de inicio y fin)."),
    ("", "4) 'Oferta': una fila por partida (código, capítulo, importe, actividad, fuente de avance, unidad y cantidad)."),
    ("", "5) 'Horas_Montaje': horas por unidad de cada partida que se monte (para la partida de montaje DERIVADO)."),
    ("", "6) En Revit, crear los parámetros compartidos CC_* (ver README) y exportar la tabla a 'Revit_Export'."),
    ("Cada semana", "a) Pegar la exportación de Revit en 'Revit_Export' (columnas A a J)."),
    ("", "b) Registrar albaranes en 'Suministros', % manual en 'Avance_Manual' y planillas en 'Planillas'."),
    ("", "c) Subir la semana de corte en 'Parametros' y guardar."),
    ("", "d) Abrir App_Control_Proyecto.html y cargar este Excel: genera el informe del directorio."),
    ("", ""),
    ("Fuente de avance", "ALBARAN = suministro, ganado al entregarse (hoja 'Suministros') · REVIT = unidades instaladas en "
                         "el modelo · MANUAL = % semanal en 'Avance_Manual' · DERIVADO = montaje, por horas de los "
                         "elementos instalados (hoja 'Horas_Montaje')."),
    ("Perfil del planificado", "ENTREGA = 100 % del importe en la semana de entrega prevista (columna X; en blanco = "
                               "inicio de la actividad) · LINEAL = repartido en la duración de la actividad."),
    ("Valor ganado (EV)", "Unidades entregadas o instaladas × precio de oferta (tope 100 % por partida). "
                          "NO es lo que certifica el instalador."),
    ("Coste real (AC)", "Lo certificado en planillas (lo que se debe). Lo pagado se controla con la SEMANA PAGO. "
                        "AC > EV = se certifica por delante de lo hecho."),
    ("Capacidad", f"Hasta {MAX_PART} partidas, {MAX_ACT} actividades, {MAX_REVIT} elementos Revit, "
                  f"{MAX_PLAN} líneas de planilla y {MAX_SEM} semanas."),
]
for i, (a, b) in enumerate(filas, 3):
    put(ws, f"A{i}", a, f_bold)
    put(ws, f"B{i}", b).alignment = Alignment(wrap_text=True, vertical="top")
ws["A4"].font = f_input
ws["A6"].font = f_link
ws["A7"].fill = fill_key

# ================================================================ Parametros
wp = wb.create_sheet("Parametros")
for col, w in zip("ABC", (34, 26, 70)):
    wp.column_dimensions[col].width = w
put(wp, "A1", "Parámetros del proyecto", f_title)
from datetime import date  # noqa: E402
param = [
    ("Proyecto", E.PROYECTO if E else "NOMBRE DEL PROYECTO", None, "Aparece en el informe."),
    ("Fecha inicio (lunes semana 1)", E.FECHA_INICIO if E else date(2027, 1, 4), "dd/mm/yyyy", "Lunes de la semana 1."),
    ("Nº semanas del cronograma", E.N_SEMANAS if E else 24, "0", f"Máximo {MAX_SEM}."),
    ("Semana de corte", E.SEMANA_CORTE if E else 1, "0", "Subir cada semana. Puede superar el nº de semanas si la obra va fuera de plazo (máx. 104)."),
    ("Fecha de corte (domingo)", "=B4+7*B6-1", "dd/mm/yyyy", "Fórmula."),
    ("Contingencia (% s/ oferta)", E.CONTINGENCIA_PCT if E else 0, PCT, "0 % hasta que dirección confirme que existe."),
    ("Umbral alerta total (proyecto)", 0.95, RATIO, "SPI/CPI del proyecto por debajo: alerta."),
    ("Umbral alerta por actividad", 0.90, RATIO, "SPI/CPI de una actividad por debajo: en rojo."),
]
for i, (a, b, fmt, nota) in enumerate(param, 3):
    put(wp, f"A{i}", a, f_bold)
    put(wp, f"B{i}", b, f_base if str(b).startswith("=") else f_input, fmt)
    put(wp, f"C{i}", nota)
for r in ("B6", "B8"):
    wp[r].fill = fill_key

# ================================================================ Cronograma
wc = wb.create_sheet("Cronograma")
cabecera(wc, ["ID_ACTIVIDAD", "ACTIVIDAD", "SEM_INICIO", "SEM_FIN", "DURACIÓN (sem)", "FECHA INICIO", "FECHA FIN",
              "PREDECESORA", "PRESUPUESTO (BAC) €", "PV a corte €", "EV a corte €", "AC a corte €", "SPI", "CPI"],
         [13, 44, 11, 10, 11, 13, 13, 13, 18, 16, 16, 16, 8, 8])
UA = MAX_ACT + 1
acts = E.ACTIVIDADES if E else []
for i in range(2, UA + 1):
    if i - 2 < len(acts):
        aid, nom, si, sf, pred = acts[i - 2]
        for col, v, fmt in (("A", aid, None), ("B", nom, None), ("C", si, "0"), ("D", sf, "0"), ("H", pred, None)):
            put(wc, f"{col}{i}", v, f_input, fmt)
    put(wc, f"E{i}", f'=IF($A{i}="","",D{i}-C{i}+1)', fmt="0")
    put(wc, f"F{i}", f'=IF($A{i}="","",Parametros!$B$4+7*(C{i}-1))', f_link, "dd/mm/yyyy")
    put(wc, f"G{i}", f'=IF($A{i}="","",Parametros!$B$4+7*D{i}-3)', f_link, "dd/mm/yyyy")
    for col, src in (("I", "F"), ("J", "N"), ("K", "R"), ("L", "S")):
        put(wc, f"{col}{i}", f'=IF($A{i}="","",SUMIFS(Oferta!${src}:${src},Oferta!$G:$G,$A{i}))', f_link, EUR0)
    put(wc, f"M{i}", f'=IF($A{i}="","",IF(N(J{i})=0,"-",K{i}/J{i}))', fmt=RATIO)
    put(wc, f"N{i}", f'=IF($A{i}="","",IF(N(L{i})=0,"-",K{i}/L{i}))', fmt=RATIO)
wc.conditional_formatting.add(
    f"M2:N{UA}", FormulaRule(formula=["AND(ISNUMBER(M2),M2<Parametros!$B$10)"],
                             fill=PatternFill("solid", fgColor="F4CCCC"), font=Font(name=F, color="9C0006")))

# ================================================================ Oferta
wo = wb.create_sheet("Oferta", 1)
cabecera(wo, ["COD_PARTIDA", "CAPÍTULO", "DESCRIPCIÓN", "CANT. OFERTA", "IMPORTE OFERTA €", "BAC €", "ID_ACTIVIDAD",
              "FUENTE AVANCE", "UD MEDICIÓN", "CANT. MEDICIÓN", "PU MEDICIÓN €", "SEM_INICIO", "SEM_FIN",
              "PV a corte €", "CANT. ENTREGADA / INSTALADA", "ELEM. INSTALADOS", "% AVANCE", "EV a corte €",
              "AC a corte €", "CV €", "SV €", "ELEM. MODELO", "PERFIL PV", "SEM. ENTREGA PREVISTA"],
         [11, 22, 40, 11, 16, 16, 12, 12, 10, 12, 13, 9, 9, 15, 14, 12, 9, 15, 15, 14, 14, 12, 11, 12])
UP = MAX_PART + 1
LASTSEM = get_column_letter(2 + MAX_SEM)
parts = E.PARTIDAS if E else []
for i in range(2, UP + 1):
    if i - 2 < len(parts):
        cod, cap, des, cant, imp, act, fuente, udm, cm = parts[i - 2]
        vals = (("A", cod, None), ("B", cap, None), ("C", des, None), ("D", cant, "#,##0"), ("E", imp, EUR),
                ("G", act, None), ("H", fuente, None), ("I", udm, None), ("J", cm, "#,##0.00"),
                ("W", "ENTREGA" if fuente == "ALBARAN" else "LINEAL", None))
        for col, v, fmt in vals:
            put(wo, f"{col}{i}", v, f_input, fmt)
    blank = f'$A{i}=""'
    put(wo, f"F{i}", f'=IF({blank},"",E{i})', fmt=EUR)
    put(wo, f"K{i}", f'=IF(OR({blank},N(J{i})=0),"",F{i}/J{i})', fmt=EUR)
    put(wo, f"L{i}", f'=IF({blank},"",IFERROR(INDEX(Cronograma!$C$2:$C${UA},MATCH(G{i},Cronograma!$A$2:$A${UA},0)),""))',
        f_link, "0")
    put(wo, f"M{i}", f'=IF({blank},"",IFERROR(INDEX(Cronograma!$D$2:$D${UA},MATCH(G{i},Cronograma!$A$2:$A${UA},0)),""))',
        f_link, "0")
    put(wo, f"N{i}", f'=IF({blank},"",IFERROR(IF(W{i}="ENTREGA",IF({CORTE}>=IF(X{i}="",L{i},X{i}),F{i},0),'
                     f'F{i}*MIN(1,MAX(0,({CORTE}-L{i}+1)/(M{i}-L{i}+1)))),0))', fmt=EUR)
    put(wo, f"O{i}", f'=IF({blank},"",IF(H{i}="REVIT",SUMIFS(Revit_Export!$G:$G,Revit_Export!$E:$E,A{i},'
                     f'Revit_Export!$I:$I,"Instalado",Revit_Export!$J:$J,"<="&{CORTE}),'
                     f'IF(H{i}="ALBARAN",SUMIFS(Suministros!$C:$C,Suministros!$B:$B,A{i},'
                     f'Suministros!$A:$A,"<="&{CORTE}),0)))', f_link, "#,##0.00")
    put(wo, f"P{i}", f'=IF({blank},"",COUNTIFS(Revit_Export!$E:$E,A{i},Revit_Export!$I:$I,"Instalado",'
                     f'Revit_Export!$J:$J,"<="&{CORTE}))', f_link, "#,##0")
    put(wo, f"Q{i}", f'=IF({blank},"",IF(OR(H{i}="REVIT",H{i}="ALBARAN"),IF(N(J{i})=0,0,MIN(1,O{i}/J{i})),'
                     f'IF(H{i}="MANUAL",MIN(1,MAX(0,IFERROR(INDEX(Avance_Manual!$C$2:${LASTSEM}${MAX_MAN + 1},'
                     f'MATCH(A{i},Avance_Manual!$A$2:$A${MAX_MAN + 1},0),{CORTE}),0))),'
                     f'IFERROR(SUM(Revit_Export!$L:$L)/SUM(Revit_Export!$K:$K),0))))', fmt=PCT)
    put(wo, f"R{i}", f'=IF({blank},"",F{i}*Q{i})', fmt=EUR)
    put(wo, f"S{i}", f'=IF({blank},"",SUMIFS(Planillas!$F:$F,Planillas!$C:$C,A{i},Planillas!$B:$B,"<="&{CORTE}))',
        f_link, EUR)
    put(wo, f"T{i}", f'=IF({blank},"",R{i}-S{i})', fmt=EUR)
    put(wo, f"U{i}", f'=IF({blank},"",R{i}-N{i})', fmt=EUR)
    put(wo, f"V{i}", f'=IF({blank},"",COUNTIFS(Revit_Export!$E:$E,A{i}))', f_link, "#,##0")
    wo[f"X{i}"].font = f_input
    wo[f"X{i}"].number_format = "0"
wo["X1"].comment = Comment("Solo para perfil ENTREGA. En blanco = semana de inicio de la actividad. "
                           "Escribir la semana de entrega pactada con el proveedor.", "Plantilla")
wo["J1"].comment = Comment("Cantidad con la que se mide el avance (ud, m, módulos...). Si la oferta es '1 ud global', "
                           "poner aquí la cantidad del modelo (p. ej. metros de tubería).", "Plantilla")
TOT = UP + 1
put(wo, f"C{TOT}", "TOTAL OFERTA", f_bold)
for col in "EFNRSTU":
    put(wo, f"{col}{TOT}", f"=SUM({col}2:{col}{UP})", f_bold, EUR)
for col in range(1, 25):
    wo.cell(row=TOT, column=col).fill = fill_total
put(wo, f"C{TOT + 1}", "Control: importe total según la oferta firmada")
c = put(wo, f"E{TOT + 1}", E.IMPORTE_CONTROL if E else None, f_input, EUR)
c.fill = fill_key
put(wo, f"F{TOT + 1}", f'=IF(E{TOT + 1}="","",IF(ABS(E{TOT}-E{TOT + 1})<0.01,"OK","REVISAR"))', f_bold)
lista(wo, f"H2:H{UP}", ["ALBARAN", "REVIT", "MANUAL", "DERIVADO"])
lista(wo, f"W2:W{UP}", ["ENTREGA", "LINEAL"])

# ================================================================ Revit_Export
wr = wb.create_sheet("Revit_Export")
cabecera(wr, ["ElementId", "Categoría", "Familia : Tipo", "Zona / Nivel", "CC_COD_PARTIDA", "CC_ID_ACTIVIDAD",
              "CANTIDAD", "UD", "CC_ESTADO", "CC_SEMANA_INSTALADO", "HORAS MONTAJE", "HORAS GANADAS"],
         [11, 22, 40, 16, 15, 15, 11, 7, 13, 18, 13, 13])
wr["A1"].comment = Comment("Pegar aquí (columnas A a J) la tabla de planificación exportada de Revit. "
                           "Las columnas K y L se calculan solas.", "Plantilla")
fila = 2
if E:
    eid = 412000
    for cod, cap, des, cant, imp, act, fuente, udm, cm in E.PARTIDAS:
        if fuente not in E.EN_MODELO:
            continue
        semanas = []
        for sem, k in sorted(E.INSTALADO.get(cod, {}).items()):
            semanas += [sem] * k
        n_el = E.N_TRAMOS if udm == "m" else int(cm)
        for j in range(n_el):
            eid += 7
            cat = E.CATEGORIA.get(cod, "Mechanical Equipment")
            if udm == "m":
                q, tipo = E.TRAMOS[j], ("Tubería acero NH3 : Tramo" if cod == "P24" else "Aislamiento coquilla : Tramo")
                tipo = f"{tipo} {j + 1:02d}"
                zona = ["Sala de máquinas", "Cámaras", "Muelles", "Frescos"][j % 4]
            else:
                q = 1
                tipo = {"Generic Models": "Soporte evaporador : Tipo A", "Pipes": "Desagüe PVC : DN40"}.get(cat, des)
                if cod == "E01":
                    tipo = f"Cuadro eléctrico CCM : Módulo {j + 1:02d}"
                zona = E.ZONAS.get(cod, "Sala de máquinas")
            estado, sem = ("Instalado", semanas[j]) if j < len(semanas) else ("Pendiente", None)
            for col, v in zip("ABCDEFGHIJ", [eid, cat, tipo, zona, cod, act, q, udm, estado, sem]):
                put(wr, f"{col}{fila}", v, f_input, "#,##0.00" if col == "G" else None)
            fila += 1
for i in range(2, MAX_REVIT + 2):
    put(wr, f"K{i}", f'=IF($E{i}="","",IFERROR(INDEX(Horas_Montaje!$C:$C,MATCH(E{i},Horas_Montaje!$A:$A,0)),0)*N(G{i}))',
        f_link, "#,##0.0")
    put(wr, f"L{i}", f'=IF($E{i}="","",IF(AND(I{i}="Instalado",ISNUMBER(J{i})),IF(J{i}<={CORTE},K{i},0),0))',
        fmt="#,##0.0")
lista(wr, f"I2:I{MAX_REVIT + 1}", ["Pendiente", "Instalado"])

# ================================================================ Suministros
wsu = wb.create_sheet("Suministros")
cabecera(wsu, ["SEMANA", "COD_PARTIDA", "CANTIDAD ENTREGADA", "Nº ALBARÁN", "OBSERVACIÓN"], [9, 13, 18, 16, 50])
if E:
    fila = 2
    for cod in sorted(E.SUMINISTRO):
        for sem, q in sorted(E.SUMINISTRO[cod].items()):
            for col, v in zip("ABCDE", [sem, cod, q, f"ALB-{sem:02d}-{cod}", ""]):
                put(wsu, f"{col}{fila}", v, f_input)
            fila += 1

# ================================================================ Horas_Montaje
wh = wb.create_sheet("Horas_Montaje")
cabecera(wh, ["COD_PARTIDA", "DESCRIPCIÓN", "HORAS POR UD", "UD"], [13, 44, 14, 8])
wh["C1"].comment = Comment("Horas-hombre de montaje por unidad (por metro en tuberías). Pedirlas al instalador.",
                           "Plantilla")
if E:
    for i, cod in enumerate(E.HORAS, 2):
        p = next(x for x in E.PARTIDAS if x[0] == cod)
        put(wh, f"A{i}", cod, f_input)
        put(wh, f"B{i}", p[2], f_input)
        put(wh, f"C{i}", E.HORAS[cod], f_input, "#,##0.0").fill = fill_key
        put(wh, f"D{i}", "h/m" if p[7] == "m" else "h/ud", f_input)

# ================================================================ Avance_Manual
wm = wb.create_sheet("Avance_Manual")
cabecera(wm, ["COD_PARTIDA", "DESCRIPCIÓN"] + [f"S{s}" for s in range(1, MAX_SEM + 1)], [12, 34] + [6] * MAX_SEM)
wm["C1"].comment = Comment("% ACUMULADO al final de cada semana (0 % a 100 %). Semanas futuras en blanco.", "Plantilla")
if E:
    r = 2
    for p in E.PARTIDAS:
        if p[6] != "MANUAL":
            continue
        put(wm, f"A{r}", p[0], f_input)
        put(wm, f"B{r}", p[2], f_input)
        for s in range(1, E.SEMANA_CORTE + 1):
            put(wm, f"{get_column_letter(2 + s)}{r}", 0, f_input, "0%")
        r += 1

# ================================================================ Planillas
wl = wb.create_sheet("Planillas")
cabecera(wl, ["Nº PLANILLA", "SEMANA", "COD_PARTIDA", "CANT. CERTIFICADA", "PRECIO APLICADO €", "IMPORTE €",
              "OBSERVACIÓN", "SEMANA PAGO"], [12, 9, 13, 16, 16, 15, 60, 12])
wl["H1"].comment = Comment("Semana en que se pagó la planilla. En blanco = pendiente de pago.", "Plantilla")
if E:
    def pu_med(cod):
        p = next(x for x in E.PARTIDAS if x[0] == cod)
        return p[4] / p[8]

    def frac_horas(s):
        hechas = total = 0.0
        for cod, h in E.HORAS.items():
            p = next(x for x in E.PARTIDAS if x[0] == cod)
            inst = E.INSTALADO.get(cod, {})
            if p[7] == "m":
                total += h * sum(E.TRAMOS)
                hechas += h * sum(E.TRAMOS[:sum(k for w, k in inst.items() if w <= s)])
            else:
                total += h * p[8]
                hechas += h * sum(k for w, k in inst.items() if w <= s)
        return hechas / total

    fila, montaje_cert = 2, 0.0
    for s in range(3, E.SEMANA_CORTE + 1):
        lineas = []
        for cod, cap, des, cant, imp, act, fuente, udm, cm in E.PARTIDAS:
            if fuente not in E.EN_MODELO:
                continue
            k = (E.SUMINISTRO if fuente == "ALBARAN" else E.INSTALADO).get(cod, {}).get(s, 0)
            if not k:
                continue
            pu, obs = round(pu_med(cod), 2), ""
            if udm == "m":
                ini = sum(v for w, v in E.INSTALADO[cod].items() if w < s)
                q = round(sum(E.TRAMOS[ini:ini + k]), 2)
                if cod == "P24":
                    q, obs = round(q * 1.08, 2), "Medición del instalador +8 % sobre metros instalados en modelo"
            else:
                q = k
            if cod == "P12":
                pu, obs = round(pu * 1.06, 2), "Precio contradictorio +6 % (pendiente de aprobar)"
            lineas.append((cod, q, pu, obs))
        if s == 9:
            lineas.append(("P17", 0.5, round(pu_med("P17"), 2), "ACOPIO 50 %: sin albarán de entrega"))
        objetivo = min(1.0, frac_horas(s) + (0.04 if s >= 5 else 0.0))
        inc = round(objetivo - montaje_cert, 4)
        if inc > 0:
            montaje_cert += inc
            lineas.append(("P25", inc, 247117.45, "Montaje: % declarado por instalador"))
        for cod, q, pu, obs in lineas:
            vals = [f"PL-{s:02d}", s, cod, q, pu, None, obs,
                    s + E.SEMANAS_PAGO if s + E.SEMANAS_PAGO <= E.SEMANA_CORTE else None]
            for col, v in zip("ABCDEFGH", vals):
                if col != "F":
                    put(wl, f"{col}{fila}", v, f_input, EUR if col == "E" else None)
            fila += 1
for i in range(2, MAX_PLAN + 2):
    put(wl, f"F{i}", f'=IF(D{i}="","",ROUND(D{i}*E{i},2))', fmt=EUR)

# ================================================================ Extras
wx = wb.create_sheet("Extras")
cabecera(wx, ["COD", "DESCRIPCIÓN", "IMPORTE €", "ESTADO", "SEMANA LÍMITE DECISIÓN", "MOTIVO DEL LÍMITE"],
         [8, 44, 15, 14, 14, 70])
UX = MAX_EXT + 1
if E:
    extras = [
        ("X01", "Cableado eléctrico receptores frío", 134200.00, "Sin decidir", 13,
         "Sin cableado de receptores no hay conexionado ni puesta en marcha. No es opcional: es alcance no contratado."),
        ("X02", "Telecontrol de la instalación en factoría", 21594.00, "Sin decidir", 17,
         "Debe integrarse antes de cerrar el cuadro."),
        ("X03", "Telecontrol vía GSM (incremento sobre X02)", 2440.00, "Sin decidir", 17, "Depende de X02."),
        ("X04", "Avisador telefónico de alarmas", 1830.00, "Sin decidir", 20,
         "Necesario antes de cargar amoniaco (seguridad)."),
    ]
    for i, row in enumerate(extras, 2):
        for col, v in zip("ABCDEF", row):
            put(wx, f"{col}{i}", v, f_input, EUR if col == "C" else None)
for col, w in (("H", 24), ("I", 16)):
    wx.column_dimensions[col].width = w
put(wx, "H2", "TOTAL EXTRAS", f_bold)
put(wx, "I2", f"=SUM(C2:C{UX})", f_bold, EUR)
put(wx, "H3", "Pendiente de decisión", f_bold)
put(wx, "I3", f'=SUMIFS(C2:C{UX},D2:D{UX},"Sin decidir")', fmt=EUR)
lista(wx, f"D2:D{UX}", ["Sin decidir", "Aprobado", "Rechazado"])

# ================================================================ Contingencia
wg = wb.create_sheet("Contingencia")
cabecera(wg, ["SEMANA", "CONCEPTO", "IMPORTE €", "ESTADO"], [9, 60, 15, 14])
if E:
    for col, v in zip("ABCD", [11, "Precio contradictorio muelles tipo 2 (+6 % s/ 13 ud)", round(48900.60 * 0.06, 2),
                               "Propuesto"]):
        put(wg, f"{col}2", v, f_input, EUR if col == "C" else None)
lista(wg, f"D2:D{MAX_CONT + 1}", ["Propuesto", "Aprobado", "Rechazado"])
for r, (a, f) in enumerate([("Contingencia inicial", f"=Parametros!$B$8*Oferta!$F${TOT}"),
                            ("Consumida (aprobado)", '=SUMIFS(C:C,D:D,"Aprobado")'),
                            ("Comprometida (propuesto)", '=SUMIFS(C:C,D:D,"Propuesto")'),
                            ("Disponible", "=G2-G3-G4")], 2):
    put(wg, f"F{r}", a, f_bold)
    put(wg, f"G{r}", f, f_link if "!" in f else f_base, EUR)
wg.column_dimensions["F"].width = 26
wg.column_dimensions["G"].width = 16

# ================================================================ Resumen
ws = wb.create_sheet("Resumen", 1)
for col, w in zip("ABC", (44, 18, 70)):
    ws.column_dimensions[col].width = w
put(ws, "A1", "Resumen a fecha de corte (valor ganado)", f_title)
put(ws, "A2", "Semana de corte", f_bold)
put(ws, "B2", "=Parametros!B6", f_link, "0")
put(ws, "A3", "Fecha de corte", f_bold)
put(ws, "B3", "=Parametros!B7", f_link, "dd/mm/yyyy")
kpi = [
    ("BAC · Presupuesto a la conclusión", f"=Oferta!F{TOT}", EUR0, "Oferta contratada."),
    ("PV · Valor planificado", f"=Oferta!N{TOT}", EUR0, "Equipos en su entrega; resto lineal en su actividad."),
    ("EV · Valor ganado", f"=Oferta!R{TOT}", EUR0, "Entregado / instalado × precio de oferta."),
    ("AC · Coste real (certificado)", f"=Oferta!S{TOT}", EUR0, "Planillas acumuladas."),
    ("SV · Variación de plazo (€)", "=B7-B6", EUR0, "EV − PV."),
    ("CV · Variación de coste (€)", "=B7-B8", EUR0, "EV − AC. Negativo = se ha certificado más de lo ganado."),
    ("SPI", '=IF(B6=0,"-",B7/B6)', RATIO, "< 1 = retraso."),
    ("CPI", '=IF(B8=0,"-",B7/B8)', RATIO, "< 1 = sobrecoste o certificación adelantada."),
    ("EAC · Estimación a la conclusión (BAC/CPI)", '=IF(AND(ISNUMBER(B12),B12<>0),B5/B12,B8+B5-B7)', EUR0,
     "Sin CPI medible: AC + (BAC − EV)."),
    ("VAC · Variación a la conclusión", "=B5-B13", EUR0, "BAC − EAC."),
    ("TCPI · Eficiencia necesaria para cumplir BAC", '=IF(B5-B7<=0.01,"-",IF(B8>=B5,"Inalcanzable",(B5-B7)/(B5-B8)))',
     RATIO, "> 1 = hay que mejorar."),
    ("% avance físico (EV/BAC)", "=IFERROR(B7/B5,0)", PCT, ""),
    ("% avance planificado (PV/BAC)", "=IFERROR(B6/B5,0)", PCT, ""),
    ("Extras sin decidir", "=Extras!I3", EUR0, "No están en el BAC."),
    ("Contingencia disponible", "=Contingencia!G5", EUR0, ""),
    ("Exposición máxima (EAC + extras sin decidir)", "=B13+B18", EUR0, "Lo que el directorio debería tener presupuestado."),
    ("Pagado a fecha de corte", f'=SUMIFS(Planillas!$F:$F,Planillas!$B:$B,"<="&{CORTE},Planillas!$H:$H,"<="&{CORTE})',
     EUR0, "Planillas certificadas y pagadas hasta el corte."),
    ("Certificado pendiente de pago", "=B8-B21", EUR0, "Deuda con el instalador a fecha de corte."),
]
for i, (a, f, fmt, nota) in enumerate(kpi, 5):
    put(ws, f"A{i}", a, f_bold)
    put(ws, f"B{i}", f, f_link if "!" in f else f_base, fmt)
    put(ws, f"C{i}", nota)
for ref in ("A13", "B13", "A20", "B20"):
    ws[ref].fill = fill_total

SALIDA.parent.mkdir(parents=True, exist_ok=True)
wb.save(SALIDA)
print(f"OK -> {SALIDA}")
