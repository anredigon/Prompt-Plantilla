"""Datos del EJEMPLO (maqueta): importes de la oferta "Costes GPT 09 0008" + avance, Revit, albaranes y planillas FICTICIOS."""
from datetime import date

PROYECTO = "Instalación frigorífica NH3 / glycol (EJEMPLO)"
FECHA_INICIO = date(2026, 9, 7)
N_SEMANAS = 24
SEMANA_CORTE = 11
CONTINGENCIA_PCT = 0.0
SEMANAS_PAGO = 6
IMPORTE_CONTROL = 1375000

# ---------------------------------------------------------------- datos
# (cod, capitulo, descripcion, cant_oferta, importe_pdf, actividad, fuente, ud_med, cant_med)
FRIO, ELEC, AIS = "Instalación frigorífica", "Electricidad", "Aislamientos"
PARTIDAS = [
    ("P01", FRIO, "MYCOM 200VLD - 270 CV", 2, 161194.95, "ACT-010", "ALBARAN", "ud", 2),
    ("P02", FRIO, "MYCOM 160VLD - 180 CV", 2, 130970.90, "ACT-010", "ALBARAN", "ud", 2),
    ("P03", FRIO, "Separador de partículas de baja", 1, 28931.41, "ACT-020", "ALBARAN", "ud", 1),
    ("P04", FRIO, "Separador de partículas de alta", 1, 20712.98, "ACT-020", "ALBARAN", "ud", 1),
    ("P05", FRIO, "Condensador evaporativo", 1, 61369.09, "ACT-030", "ALBARAN", "ud", 1),
    ("P06", FRIO, "Recipiente de líquido", 1, 12973.42, "ACT-020", "ALBARAN", "ud", 1),
    ("P07", FRIO, "Evaporador cámara 1", 2, 30291.99, "ACT-060", "ALBARAN", "ud", 2),
    ("P08", FRIO, "Evaporador cámara 2", 2, 35374.67, "ACT-060", "ALBARAN", "ud", 2),
    ("P09", FRIO, "Evaporador cámara 3", 2, 24383.19, "ACT-060", "ALBARAN", "ud", 2),
    ("P10", FRIO, "Evaporadores cámaras 4, 5 y 6", 6, 64837.95, "ACT-060", "ALBARAN", "ud", 6),
    ("P11", FRIO, "Muelles tipo 1", 17, 35883.56, "ACT-070", "ALBARAN", "ud", 17),
    ("P12", FRIO, "Muelles tipo 2", 13, 48900.60, "ACT-070", "ALBARAN", "ud", 13),
    ("P13", FRIO, "Antecámara despojos", 1, 11189.45, "ACT-080", "ALBARAN", "ud", 1),
    ("P14", FRIO, "Despojos", 1, 3297.71, "ACT-080", "ALBARAN", "ud", 1),
    ("P15", FRIO, "Amp frescos, MP & PT", 6, 28958.72, "ACT-080", "ALBARAN", "ud", 6),
    ("P16", FRIO, "Cámara de frescos", 2, 18182.39, "ACT-080", "ALBARAN", "ud", 2),
    ("P17", FRIO, "Central de bombeo de glycol 700 kW -4/-8 ºC", 1, 28328.70, "ACT-040", "ALBARAN", "ud", 1),
    ("P18", FRIO, "Protecciones de instalación", 1, 8186.29, "ACT-130", "MANUAL", "%", 1),
    ("P19", FRIO, "Desagües", 52, 9822.82, "ACT-090", "REVIT", "ud", 52),
    ("P20", FRIO, "Soportes de evaporadores", 52, 11787.38, "ACT-090", "REVIT", "ud", 52),
    ("P21", FRIO, "Cargas de aceite", 600, 4533.61, "ACT-140", "MANUAL", "l", 600),
    ("P22", FRIO, "Carga de amoniaco", 8000, 17126.96, "ACT-140", "MANUAL", "kg", 8000),
    ("P23", FRIO, "Carga de propylen glycol", 2500, 7870.85, "ACT-140", "MANUAL", "l", 2500),
    ("P24", FRIO, "Tuberías", 1, 54663.16, "ACT-050", "REVIT", "m", 1850),
    ("P25", FRIO, "Montaje de instalación de frío", 1, 247117.45, "ACT-100", "DERIVADO", "global", 1),
    ("E01", ELEC, "Cuadro eléctrico de la instalación", 1, 154799.44, "ACT-120", "ALBARAN", "módulo", 6),
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

# Suministros reales ficticios (albaranes): semana -> unidades entregadas
SUMINISTRO = {
    "P01": {4: 2}, "P02": {5: 2}, "P03": {5: 1}, "P04": {5: 1}, "P05": {8: 1}, "P06": {6: 1},
    "P07": {6: 2}, "P08": {6: 2}, "P09": {7: 2}, "P10": {7: 6}, "P11": {8: 17}, "P12": {10: 13},
    "P14": {10: 1}, "P16": {10: 2},
}
# Horas de montaje por unidad (por m en tuberías). SUPUESTO: pedir al instalador su estimación real.
HORAS = {"P01": 120, "P02": 100, "P03": 40, "P04": 40, "P05": 80, "P06": 24, "P17": 80,
         "P07": 24, "P08": 24, "P09": 24, "P10": 24, "P11": 16, "P12": 16,
         "P13": 16, "P14": 16, "P15": 16, "P16": 16, "P19": 3, "P20": 2, "P24": 1.2}
EN_MODELO = ("REVIT", "ALBARAN")


def longitudes_tramos():
    patron = [18.5, 42.0, 27.5, 35.0, 24.0, 38.5, 31.0, 29.5, 22.5, 40.0]
    base = [patron[i % len(patron)] for i in range(N_TRAMOS)]
    factor = 1850 / sum(base)
    lon = [round(x * factor, 2) for x in base]
    lon[-1] = round(1850 - sum(lon[:-1]), 2)
    return lon


TRAMOS = longitudes_tramos()
