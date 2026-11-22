# -*- coding: utf-8 -*-
"""Crea los parámetros CC_* del control de obra en el modelo activo."""
from pyrevit import revit, forms
import control_obra as co

doc = revit.doc
opciones = sorted([(co.categoria(doc, n).Name, n) for n in co.CATEGORIAS if co.categoria(doc, n) is not None])
por_defecto = ["OST_MechanicalEquipment", "OST_PipeCurves", "OST_PipeInsulations", "OST_GenericModel", "OST_ElectricalEquipment"]
elegidas = forms.SelectFromList.show(
    [o[0] for o in opciones], title="Control de obra · categorías a controlar", multiselect=True,
    button_name="Crear parámetros", default=[o[0] for o in opciones if o[1] in por_defecto])
if not elegidas:
    forms.alert("No se ha elegido ninguna categoría.", exitscript=True)
nombres = [o[1] for o in opciones if o[0] in elegidas]
try:
    vinculadas = co.crear_parametros(doc, nombres)
except Exception as e:
    forms.alert(u"No se han podido crear los parámetros:\n{}".format(e), exitscript=True)
inicio = co.inicio_obra(doc, preguntar=True)
forms.alert(u"Parámetros CC_* listos en {} categorías:\n{}\n\nInicio de obra (semana 1): {}\n\n"
            u"Siguiente paso: selecciona elementos y usa 'Asignar partida'."
            .format(len(vinculadas), u", ".join(vinculadas), inicio.strftime("%d/%m/%Y") if inicio else u"sin definir"),
            title="Control de obra")
