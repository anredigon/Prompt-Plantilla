"""
Genera el libro Excel de la MAQUETA de control de cronograma y costes.

Fuente de costes: oferta "Costes GPT 09 0008.pdf" (importes literales del PDF).
Cronograma, exportación Revit, avance y planillas: DATOS FICTICIOS para la maqueta.

Uso:  python crear_maqueta_excel.py  [ruta_salida.xlsx]
"""
import sys
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

SALIDA = Path(sys.argv[1]) if len(sys.argv) > 1 else (
    Path(__file__).resolve().parent.parent / "datos" / "Control_Proyecto_Maqueta.xlsx")

N_SEMANAS = 24
SEMANA_CORTE = 11
FECHA_INICIO = date(2026, 9, 7)  # lunes de la semana 1
CONTINGENCIA_PCT = 0.05

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
thin = Side(style="thin", color="BFBFBF")
borde = Border(bottom=thin)
EUR = '#,##0.00 €;(#,##0.00 €);-'
EUR0 = '#,##0 €;(#,##0 €);-'
PCT = '0.0%;(0.0%);-'
NUM = '#,##0.00;(#,##0.00);-'
RATIO = '0.00'

# ---------------------------------------------------------------- datos
# (cod, capitulo, descripcion, cant_oferta, importe_pdf, actividad, fuente, ud_med, cant_med)
FRIO, ELEC, AIS = "Instalación frigorífica", "Electricidad", "Aislamientos"
PARTIDAS = [
    ("P01", FRIO, "MYCOM 200VLD - 270 CV", 2, 161194.95, "ACT-010", "REVIT", "ud", 2),
    ("P02", FRIO, "MYCOM 160VLD - 180 CV", 2, 130970.90, "ACT-010", "REVIT", "ud", 2),
    ("P03", FRIO, "Separador de partículas de baja", 1, 28931.41, "ACT-020", "REVIT", "ud", 1),
    ("P04", FRIO, "Separador de partículas de alta", 1, 20712.98, "ACT-020", "REVIT", "ud", 1),
    ("P05", FRIO, "Condensador evaporativo", 1, 61369.09, "ACT-030", "REVIT", "ud", 1),
    ("P06", FRIO, "Recipiente de líquido", 1, 12973.42, "ACT-020", "REVIT", "ud", 1),
    ("P07", FRIO, "Evaporador cámara 1", 2, 30291.99, "ACT-060", "REVIT", "ud", 2),
    ("P08", FRIO, "Evaporador cámara 2", 2, 35374.67, "ACT-060", "REVIT", "ud", 2),
    ("P09", FRIO, "Evaporador cámara 3", 2, 24383.19, "ACT-060", "REVIT", "ud", 2),
    ("P10", FRIO, "Evaporadores cámaras 4, 5 y 6", 6, 64837.95, "ACT-060", "REVIT", "ud", 6),
    ("P11", FRIO, "Muelles tipo 1", 17, 35883.56, "ACT-070", "REVIT", "ud", 17),
    ("P12", FRIO, "Muelles tipo 2", 13, 48900.60, "ACT-070", "REVIT", "ud", 13),
    ("P13", FRIO, "Antecámara despojos", 1, 11189.45, "ACT-080", "REVIT", "ud", 1),
    ("P14", FRIO, "Despojos", 1, 3297.71, "ACT-080", "REVIT", "ud", 1),
    ("P15", FRIO, "Amp frescos, MP & PT", 6, 28958.72, "ACT-080", "REVIT", "ud", 6),
    ("P16", FRIO, "Cámara de frescos", 2, 18182.39, "ACT-080", "REVIT", "ud", 2),
    ("P17", FRIO, "Central de bombeo de glycol 700 kW -4/-8 ºC", 1, 28328.70, "ACT-040", "REVIT", "ud", 1),
    ("P18", FRIO, "Protecciones de instalación", 1, 8186.29, "ACT-130", "MANUAL", "%", 1),
    ("P19", FRIO, "Desagües", 52, 9822.82, "ACT-090", "REVIT", "ud", 52),
    ("P20", FRIO, "Soportes de evaporadores", 52, 11787.38, "ACT-090", "REVIT", "ud", 52),
    ("P21", FRIO, "Cargas de aceite", 600, 4533.61, "ACT-140", "MANUAL", "l", 600),
    ("P22", FRIO, "Carga de amoniaco", 8000, 17126.96, "ACT-140", "MANUAL", "kg", 8000),
    ("P23", FRIO, "Carga de propylen glycol", 2500, 7870.85, "ACT-140", "MANUAL", "l", 2500),
    ("P24", FRIO, "Tuberías", 1, 54663.16, "ACT-050", "REVIT", "m", 1850),
    ("P25", FRIO, "Montaje de instalación de frío", 1, 247117.45, "ACT-100", "DERIVADO", "global", 1),
    ("E01", ELEC, "Cuadro eléctrico de la instalación", 1, 154799.44, "ACT-120", "REVIT", "módulo", 6),
    ("E02", ELEC, "Puesta en marcha", 1, 19502.29, "ACT-150", "MANUAL", "%", 1),
    ("A01", AIS, "Aislamientos térmicos de tuberías", 1, 93808.07, "ACT-110", "REVIT", "m", 1850),
]

ACTIVIDADES = [  # (id, nombre, sem_ini, sem_fin, predecesora)
    ("ACT-010", "Suministro y colocación compresores MYCOM", 3, 7, ""),
    ("ACT-020", "Separadores y recipiente de líquido", 4, 7, ""),
    ("ACT-030", "Condensador evaporativo", 5, 8, ""),
    ("ACT-040", "Central de bombeo de glycol", 6, 9, ""),
    ("ACT-050", "Tuberías frigoríficas", 4, 16, "ACT-010"),
    ("ACT-060", "Evaporadores cámaras 1 a 6", 6, 12, ""),
    ("ACT-070", "Evaporadores muelles tipo 1 y 2", 8, 15, "ACT-060"),
    ("ACT-080", "Evaporadores despojos y frescos", 10, 15, ""),
    ("ACT-090", "Soportes y desagües de evaporadores", 6, 16, ""),
    ("ACT-100", "Montaje instalación de frío (global)", 3, 18, ""),
    ("ACT-110", "Aislamiento térmico de tuberías", 10, 19, "ACT-050"),
    ("ACT-120", "Cuadro eléctrico", 12, 19, ""),
    ("ACT-130", "Protecciones de instalación", 17, 19, "ACT-120"),
    ("ACT-140", "Cargas de aceite, amoniaco y glycol", 20, 21, "ACT-110"),
    ("ACT-150", "Puesta en marcha", 22, 24, "ACT-140"),
]

ZONAS = {"P07": "Cámara 1", "P08": "Cámara 2", "P09": "Cámara 3", "P10": "Cámaras 4-6",
         "P11": "Muelles", "P12": "Muelles", "P13": "Despojos", "P14": "Despojos",
         "P15": "Frescos", "P16": "Frescos"}
CATEGORIA = {"P01": "Mechanical Equipment", "P02": "Mechanical Equipment",
             "P03": "Mechanical Equipment", "P04": "Mechanical Equipment",
             "P05": "Mechanical Equipment", "P06": "Mechanical Equipment",
             "P17": "Mechanical Equipment", "P19": "Pipes", "P20": "Generic Models",
             "P24": "Pipes", "A01": "Pipe Insulations", "E01": "Electrical Equipment"}

# Instalaciones reales ficticias: semana -> nº de elementos instalados
INSTALADO = {
    "P01": {5: 1, 6: 1}, "P02": {6: 1, 8: 1}, "P03": {6: 1}, "P04": {7: 1},
    "P05": {9: 1}, "P06": {7: 1}, "P17": {},
    "P07": {7: 2}, "P08": {8: 1, 9: 1}, "P09": {10: 2}, "P10": {10: 1, 11: 1},
    "P11": {9: 2, 10: 3, 11: 3}, "P12": {11: 1}, "P13": {}, "P14": {11: 1},
    "P15": {}, "P16": {},
    "P19": {8: 2, 9: 4, 10: 5, 11: 5},
    "P20": {6: 2, 7: 4, 8: 4, 9: 4, 10: 5, 11: 5},
    "P24": {4: 2, 5: 3, 6: 4, 7: 4, 8: 3, 9: 4, 10: 3, 11: 4},   # tramos
    "A01": {11: 2},                                                   # tramos
    "E01": {},
}
N_TRAMOS = 60


def longitudes_tramos():
    patron = [18.5, 42.0, 27.5, 35.0, 24.0, 38.5, 31.0, 29.5, 22.5, 40.0]
    base = [patron[i % len(patron)] for i in range(N_TRAMOS)]
    factor = 1850 / sum(base)
    lon = [round(x * factor, 2) for x in base]
    lon[-1] = round(1850 - sum(lon[:-1]), 2)
    return lon


TRAMOS = longitudes_tramos()


# ---------------------------------------------------------------- helpers
def cabecera(ws, fila, titulos, anchos=None):
    for c, t in enumerate(titulos, 1):
        cell = ws.cell(row=fila, column=c, value=t)
        cell.font, cell.fill = f_head, fill_head
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[fila].height = 30
    if anchos:
        for c, a in enumerate(anchos, 1):
            ws.column_dimensions[get_column_letter(c)].width = a
    ws.freeze_panes = ws.cell(row=fila + 1, column=1)


def put(ws, ref, valor, font=f_base, fmt=None):
    cell = ws[ref]
    cell.value = valor
    cell.font = font
    if fmt:
        cell.number_format = fmt
    return cell


def pu_medicion(cod):
    p = next(x for x in PARTIDAS if x[0] == cod)
    return p[4] / p[8]


# ---------------------------------------------------------------- libro
wb = Workbook()

# ---- LEEME
ws = wb.active
ws.title = "LEEME"
ws.column_dimensions["A"].width = 26
ws.column_dimensions["B"].width = 110
put(ws, "A1", "MAQUETA · Control semanal de cronograma y costes (BIM 4D/5D)", f_title)
filas = [
    ("Estado", "MAQUETA. Costes = oferta 'Costes GPT 09 0008.pdf'. Cronograma, Revit, avance y planillas = DATOS FICTICIOS."),
    ("Texto azul", "Dato de entrada (se edita a mano o se pega de una exportación)."),
    ("Texto negro", "Fórmula. No sobrescribir."),
    ("Texto verde", "Fórmula que trae datos de otra hoja."),
    ("Relleno amarillo", "Supuesto clave a confirmar (p. ej. % de contingencia, semana de corte)."),
    ("", ""),
    ("Flujo semanal", "1) Exportar de Revit la tabla de planificación con los parámetros CC_* y pegarla en 'Revit_Export'."),
    ("", "2) Actualizar 'Avance_Manual' (partidas que no están en el modelo: cargas, protecciones, puesta en marcha)."),
    ("", "3) Registrar la planilla/certificación de la semana en 'Planillas'."),
    ("", "4) Cambiar la semana de corte en 'Parametros'. 'Oferta' y 'Resumen' se recalculan solos."),
    ("", "5) Ejecutar scripts/generar_dashboard.py → HTML del directorio + CSV para Power BI."),
    ("", ""),
    ("Fuente de avance", "REVIT = unidades instaladas contadas en el modelo · MANUAL = % semanal en 'Avance_Manual' · "
                         "DERIVADO = Montaje, avanza en proporción al nº de elementos del modelo de frío ya montados (no por €: los equipos caros inflarían el avance)."),
    ("Valor ganado (EV)", "Unidades instaladas y verificadas × precio unitario de oferta. NO es lo que certifica el instalador."),
    ("Coste real (AC)", "Lo certificado en planillas. Si AC > EV se está pagando por delante de lo instalado (acopios, "
                        "sobremediciones, precios contradictorios)."),
    ("Ejemplo de fila", "Revit_Export fila 2: un compresor con CC_COD_PARTIDA=P01, CC_ESTADO=Instalado, CC_SEMANA_INSTALADO=5."),
]
for i, (a, b) in enumerate(filas, 3):
    put(ws, f"A{i}", a, f_bold)
    put(ws, f"B{i}", b).alignment = Alignment(wrap_text=True, vertical="top")
ws["A4"].font = f_input
ws["A6"].font = f_link
ws["A7"].fill = fill_key

# ---- Parametros
wp = wb.create_sheet("Parametros")
wp.column_dimensions["A"].width = 34
wp.column_dimensions["B"].width = 22
wp.column_dimensions["C"].width = 70
put(wp, "A1", "Parámetros del control", f_title)
param = [
    ("Proyecto", "Instalación frigorífica NH3 / glycol", None, "Dato del usuario."),
    ("Fecha inicio (lunes semana 1)", FECHA_INICIO, "dd/mm/yyyy", "Supuesto de la maqueta."),
    ("Nº semanas del cronograma", N_SEMANAS, "0", "Supuesto de la maqueta."),
    ("Semana de corte", SEMANA_CORTE, "0", "Cambiar cada semana."),
    ("Fecha de corte (domingo)", "=B4+7*B6-1", "dd/mm/yyyy", "Fórmula."),
    ("Contingencia (% s/ oferta)", CONTINGENCIA_PCT, PCT, "SUPUESTO: no consta en la oferta. Confirmar con dirección."),
    ("Umbral alerta SPI / CPI", 0.95, RATIO, "Por debajo de este valor la actividad se marca en rojo."),
]
for i, (a, b, fmt, nota) in enumerate(param, 3):
    put(wp, f"A{i}", a, f_bold)
    c = put(wp, f"B{i}", b, f_base if str(b).startswith("=") else f_input, fmt)
    put(wp, f"C{i}", nota)
for r in ("B6", "B8"):
    wp[r].fill = fill_key
CORTE = "Parametros!$B$6"

# ---- Cronograma
wc = wb.create_sheet("Cronograma")
cabecera(wc, 1, ["ID_ACTIVIDAD", "ACTIVIDAD", "SEM_INICIO", "SEM_FIN", "DURACIÓN (sem)",
                 "FECHA INICIO", "FECHA FIN", "PREDECESORA", "PRESUPUESTO (BAC) €",
                 "PV a corte €", "EV a corte €", "AC a corte €", "SPI", "CPI"],
         [13, 44, 11, 10, 11, 13, 13, 13, 18, 16, 16, 16, 8, 8])
for i, (aid, nom, si, sf, pred) in enumerate(ACTIVIDADES, 2):
    put(wc, f"A{i}", aid, f_input)
    put(wc, f"B{i}", nom, f_input)
    put(wc, f"C{i}", si, f_input, "0")
    put(wc, f"D{i}", sf, f_input, "0")
    put(wc, f"E{i}", f"=D{i}-C{i}+1", fmt="0")
    put(wc, f"F{i}", f"=Parametros!$B$4+7*(C{i}-1)", f_link, "dd/mm/yyyy")
    put(wc, f"G{i}", f"=Parametros!$B$4+7*D{i}-3", f_link, "dd/mm/yyyy")
    put(wc, f"H{i}", pred, f_input)
    put(wc, f"I{i}", f"=SUMIFS(Oferta!$F:$F,Oferta!$G:$G,A{i})", f_link, EUR0)
    put(wc, f"J{i}", f"=SUMIFS(Oferta!$N:$N,Oferta!$G:$G,A{i})", f_link, EUR0)
    put(wc, f"K{i}", f"=SUMIFS(Oferta!$R:$R,Oferta!$G:$G,A{i})", f_link, EUR0)
    put(wc, f"L{i}", f"=SUMIFS(Oferta!$S:$S,Oferta!$G:$G,A{i})", f_link, EUR0)
    put(wc, f"M{i}", f'=IF(J{i}=0,"-",K{i}/J{i})', fmt=RATIO)
    put(wc, f"N{i}", f'=IF(L{i}=0,"-",K{i}/L{i})', fmt=RATIO)
nA = len(ACTIVIDADES) + 1
wc.conditional_formatting.add(
    f"M2:N{nA}", FormulaRule(formula=["AND(ISNUMBER(M2),M2<Parametros!$B$9)"],
                             fill=PatternFill("solid", fgColor="F4CCCC"), font=Font(name=F, color="9C0006")))

# ---- Oferta (partidas)
wo = wb.create_sheet("Oferta", 1)
cabecera(wo, 1, ["COD_PARTIDA", "CAPÍTULO", "DESCRIPCIÓN", "CANT. OFERTA", "IMPORTE OFERTA €",
                 "BAC €", "ID_ACTIVIDAD", "FUENTE AVANCE", "UD MEDICIÓN", "CANT. MEDICIÓN",
                 "PU MEDICIÓN €", "SEM_INICIO", "SEM_FIN", "PV a corte €", "CANT. INSTALADA",
                 "ELEM. INSTALADOS", "% AVANCE", "EV a corte €", "AC a corte €", "CV €", "SV €", "ELEM. MODELO"],
         [11, 22, 40, 11, 16, 16, 12, 12, 10, 12, 13, 9, 9, 15, 12, 12, 9, 15, 15, 14, 14, 12])
n = len(PARTIDAS)
UP = n + 1
for i, (cod, cap, des, cant, imp, act, fuente, udm, cm) in enumerate(PARTIDAS, 2):
    put(wo, f"A{i}", cod, f_input)
    put(wo, f"B{i}", cap, f_input)
    put(wo, f"C{i}", des, f_input)
    put(wo, f"D{i}", cant, f_input, "#,##0")
    put(wo, f"E{i}", imp, f_input, EUR)
    put(wo, f"F{i}", f"=E{i}", fmt=EUR)
    put(wo, f"G{i}", act, f_input)
    put(wo, f"H{i}", fuente, f_input)
    put(wo, f"I{i}", udm, f_input)
    put(wo, f"J{i}", cm, f_input, "#,##0.00")
    put(wo, f"K{i}", f"=F{i}/J{i}", fmt=EUR)
    put(wo, f"L{i}", f"=INDEX(Cronograma!$C$2:$C${nA},MATCH(G{i},Cronograma!$A$2:$A${nA},0))", f_link, "0")
    put(wo, f"M{i}", f"=INDEX(Cronograma!$D$2:$D${nA},MATCH(G{i},Cronograma!$A$2:$A${nA},0))", f_link, "0")
    put(wo, f"N{i}", f"=F{i}*MIN(1,MAX(0,({CORTE}-L{i}+1)/(M{i}-L{i}+1)))", fmt=EUR)
    put(wo, f"O{i}", f'=IF(H{i}="REVIT",SUMIFS(Revit_Export!$G:$G,Revit_Export!$E:$E,A{i},'
                     f'Revit_Export!$I:$I,"Instalado",Revit_Export!$J:$J,"<="&{CORTE}),0)', f_link, "#,##0.00")
    put(wo, f"P{i}", f'=COUNTIFS(Revit_Export!$E:$E,A{i},Revit_Export!$I:$I,"Instalado",'
                     f'Revit_Export!$J:$J,"<="&{CORTE})', f_link, "#,##0")
    put(wo, f"V{i}", f'=COUNTIFS(Revit_Export!$E:$E,A{i})', f_link, "#,##0")
    put(wo, f"Q{i}",
        f'=IF(H{i}="REVIT",O{i}/J{i},IF(H{i}="MANUAL",'
        f'IFERROR(INDEX(Avance_Manual!$C$2:$Z$20,MATCH(A{i},Avance_Manual!$A$2:$A$20,0),{CORTE}),0),'
        f'SUMIFS($P$2:$P${UP},$B$2:$B${UP},B{i},$H$2:$H${UP},"REVIT")/'
        f'SUMIFS($V$2:$V${UP},$B$2:$B${UP},B{i},$H$2:$H${UP},"REVIT")))', fmt=PCT)
    put(wo, f"R{i}", f"=F{i}*Q{i}", fmt=EUR)
    put(wo, f"S{i}", f'=SUMIFS(Planillas!$F:$F,Planillas!$C:$C,A{i},Planillas!$B:$B,"<="&{CORTE})', f_link, EUR)
    put(wo, f"T{i}", f"=R{i}-S{i}", fmt=EUR)
    put(wo, f"U{i}", f"=R{i}-N{i}", fmt=EUR)
    if cod in ("P24", "A01", "E01"):
        wo[f"J{i}"].comment = Comment("Oferta en 1 ud global. Cantidad de medición = la que da el modelo Revit "
                                      "(supuesto maqueta). PU medición = importe / cantidad.", "Maqueta")
wo["E2"].comment = Comment("Importes tomados literalmente del PDF 'Costes GPT 09 0008' (columna TOTAL).", "Maqueta")
tot = UP + 1
put(wo, f"C{tot}", "TOTAL OFERTA", f_bold)
for col in "EFNRSTU":
    c = put(wo, f"{col}{tot}", f"=SUM({col}2:{col}{UP})", f_bold, EUR)
for col in range(1, 23):
    wo.cell(row=tot, column=col).fill = fill_total
put(wo, f"C{tot+1}", "Control: total oferta PDF")
put(wo, f"E{tot+1}", 1375000, f_input, EUR)
put(wo, f"F{tot+1}", f'=IF(ABS(E{tot}-E{tot+1})<0.01,"OK","REVISAR")', f_bold)
dv = DataValidation(type="list", formula1='"REVIT,MANUAL,DERIVADO"', allow_blank=False)
wo.add_data_validation(dv)
dv.add(f"H2:H{UP}")

# ---- Revit_Export
wr = wb.create_sheet("Revit_Export")
cabecera(wr, 1, ["ElementId", "Categoría", "Familia : Tipo", "Zona / Nivel", "CC_COD_PARTIDA",
                 "CC_ID_ACTIVIDAD", "CANTIDAD", "UD", "CC_ESTADO", "CC_SEMANA_INSTALADO"],
         [11, 22, 40, 16, 15, 15, 11, 7, 13, 18])
fila, eid = 2, 412000
for cod, cap, des, cant, imp, act, fuente, udm, cm in PARTIDAS:
    if fuente != "REVIT":
        continue
    semanas = []
    for sem, k in sorted(INSTALADO.get(cod, {}).items()):
        semanas += [sem] * k
    n_el = N_TRAMOS if udm == "m" else int(cm)
    for j in range(n_el):
        eid += 7
        if udm == "m":
            q, tipo = TRAMOS[j], ("Tubería acero NH3 : Tramo" if cod == "P24" else "Aislamiento coquilla : Tramo")
        else:
            q = 1
            tipo = {"Generic Models": "Soporte evaporador : Tipo A", "Pipes": "Desagüe PVC : DN40"}.get(
                CATEGORIA.get(cod, "Mechanical Equipment"), des)
            if cod == "E01":
                tipo = "Cuadro eléctrico CCM : Módulo"
        cat = CATEGORIA.get(cod, "Mechanical Equipment")
        estado, sem = ("Instalado", semanas[j]) if j < len(semanas) else ("Pendiente", None)
        zona = ZONAS.get(cod, "Sala de máquinas")
        if udm == "m":
            zona = ["Sala de máquinas", "Cámaras", "Muelles", "Frescos"][j % 4]
        vals = [eid, cat, f"{tipo} {j+1:02d}" if udm == "m" or cod == "E01" else tipo, zona, cod, act,
                q, udm, estado, sem]
        for c, v in enumerate(vals, 1):
            cell = wr.cell(row=fila, column=c, value=v)
            cell.font = f_input
        wr.cell(row=fila, column=7).number_format = "#,##0.00"
        fila += 1
dv2 = DataValidation(type="list", formula1='"Pendiente,Instalado"', allow_blank=False)
wr.add_data_validation(dv2)
dv2.add(f"I2:I{fila-1}")

# ---- Avance_Manual
wm = wb.create_sheet("Avance_Manual")
cabecera(wm, 1, ["COD_PARTIDA", "DESCRIPCIÓN"] + [f"S{s}" for s in range(1, N_SEMANAS + 1)],
         [12, 34] + [6] * N_SEMANAS)
put(wm, "A22", "Introducir el % ACUMULADO al final de cada semana (0 % a 100 %). Columnas futuras en blanco.")
r = 2
for cod, cap, des, *_rest in PARTIDAS:
    if _rest[3] != "MANUAL":
        continue
    put(wm, f"A{r}", cod, f_input)
    put(wm, f"B{r}", des, f_input)
    for s in range(1, SEMANA_CORTE + 1):
        put(wm, f"{get_column_letter(2 + s)}{r}", 0, f_input, "0%")
    r += 1

# ---- Planillas
wl = wb.create_sheet("Planillas")
cabecera(wl, 1, ["Nº PLANILLA", "SEMANA", "COD_PARTIDA", "CANT. CERTIFICADA", "PRECIO APLICADO €",
                 "IMPORTE €", "OBSERVACIÓN"], [12, 9, 13, 16, 16, 15, 60])


def ev_fisico_frio_semana(s):
    """Fracción de elementos del modelo de frío (REVIT) montados hasta la semana s."""
    hechos = total = 0
    for cod, cap, des, cant, imp, act, fuente, udm, cm in PARTIDAS:
        if cap != FRIO or fuente != "REVIT":
            continue
        total += N_TRAMOS if udm == "m" else int(cm)
        hechos += sum(k for w, k in INSTALADO.get(cod, {}).items() if w <= s)
    return hechos / total


fila = 2
montaje_cert = 0.0
for s in range(3, SEMANA_CORTE + 1):
    lineas = []
    for cod, cap, des, cant, imp, act, fuente, udm, cm in PARTIDAS:
        if fuente != "REVIT":
            continue
        k = INSTALADO.get(cod, {}).get(s, 0)
        if not k:
            continue
        pu = round(pu_medicion(cod), 2)
        obs = ""
        if udm == "m":
            ini = sum(v for w, v in INSTALADO[cod].items() if w < s)
            q = round(sum(TRAMOS[ini:ini + k]), 2)
            if cod == "P24":
                q = round(q * 1.08, 2)
                obs = "Medición del instalador +8 % sobre metros instalados en modelo"
        else:
            q = k
        if cod == "P12":
            pu = round(pu * 1.06, 2)
            obs = "Precio contradictorio +6 % (pendiente de aprobar)"
        lineas.append((cod, q, pu, obs))
    if s == 9:
        lineas.append(("P17", 0.5, round(pu_medicion("P17"), 2), "ACOPIO 50 %: equipo no instalado"))
    # Montaje: el instalador certifica 4 puntos por delante del avance físico
    objetivo = min(1.0, ev_fisico_frio_semana(s) + (0.04 if s >= 5 else 0.0))
    inc = round(objetivo - montaje_cert, 4)
    if inc > 0:
        montaje_cert += inc
        lineas.append(("P25", inc, 247117.45, "Montaje: % declarado por instalador"))
    for cod, q, pu, obs in lineas:
        put(wl, f"A{fila}", f"PL-{s:02d}", f_input)
        put(wl, f"B{fila}", s, f_input, "0")
        put(wl, f"C{fila}", cod, f_input)
        put(wl, f"D{fila}", q, f_input, "#,##0.00##")
        put(wl, f"E{fila}", pu, f_input, EUR)
        put(wl, f"F{fila}", f"=ROUND(D{fila}*E{fila},2)", fmt=EUR)
        put(wl, f"G{fila}", obs, f_input)
        fila += 1

# ---- Extras
wx = wb.create_sheet("Extras")
cabecera(wx, 1, ["COD", "DESCRIPCIÓN", "IMPORTE €", "ESTADO", "SEMANA LÍMITE DECISIÓN", "MOTIVO DEL LÍMITE"],
         [8, 44, 15, 14, 14, 70])
extras = [
    ("X01", "Cableado eléctrico receptores frío", 134200.00, "Sin decidir", 13,
     "Sin cableado de receptores no hay conexionado ni puesta en marcha. No es opcional: es alcance no contratado."),
    ("X02", "Telecontrol de la instalación en factoría", 21594.00, "Sin decidir", 17, "Debe integrarse antes de cerrar el cuadro."),
    ("X03", "Telecontrol vía GSM (incremento sobre X02)", 2440.00, "Sin decidir", 17, "Depende de X02."),
    ("X04", "Avisador telefónico de alarmas", 1830.00, "Sin decidir", 20, "Necesario antes de cargar amoniaco (seguridad)."),
]
for i, row in enumerate(extras, 2):
    for c, v in enumerate(row, 1):
        put(wx, f"{get_column_letter(c)}{i}", v, f_input, EUR if c == 3 else ("0" if c == 5 else None))
put(wx, "B7", "TOTAL EXTRAS", f_bold)
put(wx, "C7", "=SUM(C2:C5)", f_bold, EUR)
put(wx, "B8", "Pendiente de decisión")
put(wx, "C8", '=SUMIFS(C2:C5,D2:D5,"Sin decidir")', fmt=EUR)
dv3 = DataValidation(type="list", formula1='"Sin decidir,Aprobado,Rechazado"', allow_blank=False)
wx.add_data_validation(dv3)
dv3.add("D2:D5")
put(wx, "A10", "Importes del PDF (bloque 'SUMINISTROS EXTRA (NO INCLUIDOS)'). Semana límite y motivo: criterio de la maqueta.")

# ---- Contingencia
wg = wb.create_sheet("Contingencia")
cabecera(wg, 1, ["SEMANA", "CONCEPTO", "IMPORTE €", "ESTADO"], [9, 60, 15, 14])
put(wg, "A2", 11, f_input, "0")
put(wg, "B2", "Precio contradictorio muelles tipo 2 (+6 % s/ 13 ud)", f_input)
put(wg, "C2", round(48900.60 * 0.06, 2), f_input, EUR)
put(wg, "D2", "Propuesto", f_input)
dv4 = DataValidation(type="list", formula1='"Propuesto,Aprobado,Rechazado"', allow_blank=False)
wg.add_data_validation(dv4)
dv4.add("D2:D50")
put(wg, "F2", "Contingencia inicial", f_bold)
put(wg, "G2", "=Parametros!$B$8*Oferta!$F$" + str(tot), f_link, EUR)
put(wg, "F3", "Consumida (aprobado)", f_bold)
put(wg, "G3", '=SUMIFS(C:C,D:D,"Aprobado")', fmt=EUR)
put(wg, "F4", "Comprometida (propuesto)", f_bold)
put(wg, "G4", '=SUMIFS(C:C,D:D,"Propuesto")', fmt=EUR)
put(wg, "F5", "Disponible", f_bold)
put(wg, "G5", "=G2-G3-G4", f_bold, EUR)
wg.column_dimensions["F"].width = 26
wg.column_dimensions["G"].width = 16

# ---- Resumen
ws = wb.create_sheet("Resumen", 1)
ws.column_dimensions["A"].width = 44
ws.column_dimensions["B"].width = 18
ws.column_dimensions["C"].width = 70
put(ws, "A1", "Resumen a fecha de corte (valor ganado)", f_title)
put(ws, "A2", "Semana de corte", f_bold)
put(ws, "B2", "=Parametros!B6", f_link, "0")
put(ws, "A3", "Fecha de corte", f_bold)
put(ws, "B3", "=Parametros!B7", f_link, "dd/mm/yyyy")
O = f"Oferta!$F${tot}"
kpi = [
    ("BAC · Presupuesto a la conclusión", f"=Oferta!F{tot}", EUR0, "Oferta contratada."),
    ("PV · Valor planificado", f"=Oferta!N{tot}", EUR0, "Reparto lineal del BAC de cada partida en la duración de su actividad."),
    ("EV · Valor ganado", f"=Oferta!R{tot}", EUR0, "Unidades instaladas (Revit/manual) × precio de oferta."),
    ("AC · Coste real (certificado)", f"=Oferta!S{tot}", EUR0, "Planillas acumuladas."),
    ("SV · Variación de plazo (€)", "=B7-B6", EUR0, "EV − PV."),
    ("CV · Variación de coste (€)", "=B7-B8", EUR0, "EV − AC. Negativo = se ha pagado más de lo ganado."),
    ("SPI", '=IF(B6=0,"-",B7/B6)', RATIO, "< 1 = retraso."),
    ("CPI", '=IF(B8=0,"-",B7/B8)', RATIO, "< 1 = sobrecoste o pago adelantado."),
    ("EAC · Estimación a la conclusión (BAC/CPI)", '=IF(AND(ISNUMBER(B12),B12<>0),B5/B12,B8+B5-B7)', EUR0,
     "Hipótesis: el rendimiento de coste actual se mantiene. Sin CPI medible: AC + (BAC − EV)."),
    ("VAC · Variación a la conclusión", "=B5-B13", EUR0, "BAC − EAC."),
    ("TCPI · Eficiencia necesaria para cumplir BAC", '=IF(B5=B8,"-",(B5-B7)/(B5-B8))', RATIO, "> 1 = hay que mejorar."),
    ("% avance físico (EV/BAC)", "=IFERROR(B7/B5,0)", PCT, ""),
    ("% avance planificado (PV/BAC)", "=IFERROR(B6/B5,0)", PCT, ""),
    ("Extras sin decidir", "=Extras!C8", EUR0, "No están en el BAC."),
    ("Contingencia disponible", "=Contingencia!G5", EUR0, ""),
    ("Exposición máxima (EAC + extras sin decidir)", "=B13+B18", EUR0, "Lo que el directorio debería tener presupuestado."),
]
for i, (a, f, fmt, nota) in enumerate(kpi, 5):
    put(ws, f"A{i}", a, f_bold)
    put(ws, f"B{i}", f, f_link if "!" in f else f_base, fmt)
    put(ws, f"C{i}", nota)
for col_row in ("A13", "B13", "A20", "B20"):
    ws[col_row].fill = fill_total

for w in wb.worksheets:
    for row in w.iter_rows():
        for cell in row:
            if cell.font is None or cell.font.name != F:
                cell.font = Font(name=F, size=10)

SALIDA.parent.mkdir(parents=True, exist_ok=True)
wb.save(SALIDA)
print(f"OK -> {SALIDA}")
