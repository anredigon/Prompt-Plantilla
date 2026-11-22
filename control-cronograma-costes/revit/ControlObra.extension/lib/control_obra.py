# -*- coding: utf-8 -*-
"""Funciones comunes de la extensión Control de obra (pyRevit).

Compatible con IronPython 2.7 y CPython 3 (motores de pyRevit) y con Revit 2020-2025.
Parámetros que gestiona:
  - De ejemplar, en los elementos que se controlan: CC_COD_PARTIDA, CC_ID_ACTIVIDAD, CC_ESTADO,
    CC_SEMANA_INSTALADO, CC_ZONA.
  - En Información de proyecto: CC_INICIO_OBRA (AAAA-MM-DD, lunes de la semana 1) y CC_CARPETA_EXPORT.
"""
import datetime
import io
import os

from pyrevit import DB, forms

PARAMS_ELEMENTO = [("CC_COD_PARTIDA", "texto"), ("CC_ID_ACTIVIDAD", "texto"), ("CC_ESTADO", "texto"),
                   ("CC_SEMANA_INSTALADO", "entero"), ("CC_ZONA", "texto")]
PARAMS_PROYECTO = [("CC_INICIO_OBRA", "texto"), ("CC_CARPETA_EXPORT", "texto")]

# Categorías que se ofrecen por defecto (el usuario elige cuáles en "Preparar modelo")
CATEGORIAS = [
    "OST_MechanicalEquipment", "OST_PipeCurves", "OST_PipeInsulations", "OST_PipeFitting", "OST_PipeAccessory",
    "OST_FlexPipeCurves", "OST_PlumbingFixtures", "OST_Sprinklers", "OST_DuctCurves", "OST_DuctInsulations",
    "OST_DuctFitting", "OST_DuctAccessory", "OST_DuctTerminal", "OST_FlexDuctCurves", "OST_ElectricalEquipment",
    "OST_ElectricalFixtures", "OST_LightingFixtures", "OST_Conduit", "OST_CableTray", "OST_FireAlarmDevices",
    "OST_GenericModel", "OST_SpecialityEquipment", "OST_StructuralColumns", "OST_StructuralFraming",
    "OST_StructuralFoundation", "OST_Floors", "OST_Walls", "OST_Roofs", "OST_Ceilings", "OST_Doors",
    "OST_Windows", "OST_Stairs", "OST_Railings", "OST_Furniture",
]
# Categorías que se miden en metros de longitud (el resto, en unidades)
LINEALES = ["OST_PipeCurves", "OST_PipeInsulations", "OST_FlexPipeCurves", "OST_DuctCurves",
            "OST_DuctInsulations", "OST_FlexDuctCurves", "OST_Conduit", "OST_CableTray"]
PIE_A_METRO = 0.3048  # Revit guarda las longitudes en pies internamente
CABECERA_CSV = ["ElementId", "Categoría", "Familia : Tipo", "Zona / Nivel", "CC_COD_PARTIDA", "CC_ID_ACTIVIDAD",
                "CANTIDAD", "UD", "CC_ESTADO", "CC_SEMANA_INSTALADO"]


# ------------------------------------------------------------------ utilidades Revit
def id_entero(elem_id):
    """ElementId -> entero (IntegerValue hasta 2023, Value desde 2024)."""
    try:
        return elem_id.Value
    except AttributeError:
        return elem_id.IntegerValue


def categoria(doc, nombre_ost):
    try:
        return doc.Settings.Categories.get_Item(getattr(DB.BuiltInCategory, nombre_ost))
    except Exception:
        return None


def es_categoria(elem, nombres_ost):
    if elem.Category is None:
        return False
    cid = id_entero(elem.Category.Id)
    for nombre in nombres_ost:
        bic = getattr(DB.BuiltInCategory, nombre, None)
        if bic is not None and cid == int(bic):
            return True
    return False


def texto(elem, nombre):
    p = elem.LookupParameter(nombre)
    if p is None or not p.HasValue:
        return u""
    v = p.AsString()
    if v is None:
        v = p.AsValueString() or u""
    return v.strip()


def entero(elem, nombre):
    p = elem.LookupParameter(nombre)
    if p is None or not p.HasValue:
        return None
    try:
        return p.AsInteger()
    except Exception:
        return None


def escribir(elem, nombre, valor):
    """Escribe un parámetro. Devuelve False si el elemento no lo tiene o es de solo lectura."""
    p = elem.LookupParameter(nombre)
    if p is None or p.IsReadOnly:
        return False
    p.Set(valor)
    return True


# ------------------------------------------------------------------ parámetros compartidos
def _opciones(nombre, tipo):
    try:  # Revit 2022 y posteriores
        spec = DB.SpecTypeId.Int.Integer if tipo == "entero" else DB.SpecTypeId.String.Text
        return DB.ExternalDefinitionCreationOptions(nombre, spec)
    except AttributeError:  # Revit 2021 y anteriores
        ptype = DB.ParameterType.Integer if tipo == "entero" else DB.ParameterType.Text
        return DB.ExternalDefinitionCreationOptions(nombre, ptype)


def _vincular(doc, definicion, enlace):
    mapa = doc.ParameterBindings
    try:  # Revit 2024 y posteriores
        grupo = DB.GroupTypeId.Data
        if not mapa.Insert(definicion, enlace, grupo):
            mapa.ReInsert(definicion, enlace, grupo)
    except AttributeError:
        grupo = DB.BuiltInParameterGroup.PG_DATA
        if not mapa.Insert(definicion, enlace, grupo):
            mapa.ReInsert(definicion, enlace, grupo)


def crear_parametros(doc, nombres_ost):
    """Crea (o amplía a nuevas categorías) los parámetros CC_*. Devuelve las categorías vinculadas."""
    app = doc.Application
    original = app.SharedParametersFilename
    ruta = os.path.join(os.environ.get("TEMP") or os.path.expanduser("~"), "ControlObra_parametros_compartidos.txt")
    if not os.path.exists(ruta):
        io.open(ruta, "w", encoding="utf-8").close()
    try:
        app.SharedParametersFilename = ruta
        archivo = app.OpenSharedParameterFile()
        grupo = archivo.Groups.get_Item("ControlObra") or archivo.Groups.Create("ControlObra")
        definiciones = {}
        for nombre, tipo in PARAMS_ELEMENTO + PARAMS_PROYECTO:
            d = grupo.Definitions.get_Item(nombre)
            definiciones[nombre] = d if d is not None else grupo.Definitions.Create(_opciones(nombre, tipo))
        cats = app.Create.NewCategorySet()
        vinculadas = []
        for nombre in nombres_ost:
            c = categoria(doc, nombre)
            if c is not None and c.AllowsBoundParameters:
                cats.Insert(c)
                vinculadas.append(c.Name)
        cats_proyecto = app.Create.NewCategorySet()
        cats_proyecto.Insert(categoria(doc, "OST_ProjectInformation"))
        t = DB.Transaction(doc, "Control de obra: crear parámetros CC_")
        t.Start()
        try:
            for nombre, _ in PARAMS_ELEMENTO:
                _vincular(doc, definiciones[nombre], app.Create.NewInstanceBinding(cats))
            for nombre, _ in PARAMS_PROYECTO:
                _vincular(doc, definiciones[nombre], app.Create.NewInstanceBinding(cats_proyecto))
            t.Commit()
        except Exception:
            t.RollBack()
            raise
        return vinculadas
    finally:
        app.SharedParametersFilename = original or ""


# ------------------------------------------------------------------ datos del proyecto
def inicio_obra(doc, preguntar=True):
    """Fecha (datetime.date) del lunes de la semana 1, leída de CC_INICIO_OBRA o preguntada al usuario."""
    info = doc.ProjectInformation
    valor = texto(info, "CC_INICIO_OBRA")
    try:
        return datetime.datetime.strptime(valor, "%Y-%m-%d").date()
    except ValueError:
        pass
    if not preguntar:
        return None
    valor = forms.ask_for_string(default=datetime.date.today().strftime("%Y-%m-%d"),
                                 prompt="Fecha de inicio de obra (lunes de la semana 1), formato AAAA-MM-DD:",
                                 title="Control de obra")
    if not valor:
        return None
    try:
        fecha = datetime.datetime.strptime(valor.strip(), "%Y-%m-%d").date()
    except ValueError:
        forms.alert("La fecha debe tener el formato AAAA-MM-DD (por ejemplo 2026-09-07).", exitscript=True)
    fecha = fecha - datetime.timedelta(days=fecha.weekday())  # lunes de esa semana
    t = DB.Transaction(doc, "Control de obra: fecha de inicio")
    t.Start()
    ok = escribir(info, "CC_INICIO_OBRA", fecha.strftime("%Y-%m-%d"))
    t.Commit()
    if not ok:
        forms.alert("El modelo no tiene el parámetro CC_INICIO_OBRA. Pulsa primero 'Preparar modelo'.", exitscript=True)
    return fecha


def semana_actual(inicio, dia=None):
    dia = dia or datetime.date.today()
    return (dia - inicio).days // 7 + 1


def cantidad(elem):
    """(cantidad, unidad): metros para elementos lineales, 1 ud para el resto."""
    if es_categoria(elem, LINEALES):
        p = elem.get_Parameter(DB.BuiltInParameter.CURVE_ELEM_LENGTH)
        if p is None or not p.HasValue:
            p = elem.LookupParameter("Length") or elem.LookupParameter("Longitud")
        if p is not None and p.HasValue:
            return round(p.AsDouble() * PIE_A_METRO, 3), u"m"
    return 1, u"ud"


def familia_tipo(doc, elem):
    tipo = doc.GetElement(elem.GetTypeId())
    familia = u""
    if tipo is not None:
        try:
            familia = tipo.FamilyName
        except Exception:
            familia = u""
        try:  # en IronPython, Name de los tipos de familia solo se lee así
            nombre_tipo = DB.Element.Name.__get__(tipo)
        except Exception:
            nombre_tipo = tipo.Name
        return u"{} : {}".format(familia, nombre_tipo) if familia else nombre_tipo
    return elem.Name


def zona(doc, elem):
    z = texto(elem, "CC_ZONA")
    if z:
        return z
    try:
        nivel = doc.GetElement(elem.LevelId)
        return nivel.Name if nivel is not None else u""
    except Exception:
        return u""


def csv_linea(valores):
    salida = []
    for v in valores:
        v = u"" if v is None else (v if isinstance(v, type(u"")) else u"{}".format(v))
        if any(c in v for c in u',"\n\r'):
            v = u'"' + v.replace(u'"', u'""') + u'"'
        salida.append(v)
    return u",".join(salida) + u"\n"
