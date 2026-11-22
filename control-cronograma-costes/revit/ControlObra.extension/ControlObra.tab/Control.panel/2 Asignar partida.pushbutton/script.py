# -*- coding: utf-8 -*-
"""Asigna código de partida (y opcionalmente actividad y zona) a los elementos seleccionados."""
from pyrevit import revit, forms, DB
import control_obra as co

doc = revit.doc
elementos = revit.get_selection().elements
if not elementos:
    forms.alert("Selecciona primero los elementos (puedes usar 'Seleccionar todos los ejemplares' del tipo).", exitscript=True)
partida = forms.ask_for_string(default=co.texto(elementos[0], "CC_COD_PARTIDA"),
                               prompt=u"Código de partida del presupuesto para {} elemento(s) (p. ej. P11):".format(len(elementos)),
                               title="Control de obra · asignar partida")
if not partida:
    forms.alert("Sin código de partida no se cambia nada.", exitscript=True)
actividad = forms.ask_for_string(default=co.texto(elementos[0], "CC_ID_ACTIVIDAD"),
                                 prompt="Actividad del cronograma (opcional, p. ej. ACT-070):", title="Control de obra") or ""
zona = forms.ask_for_string(default=co.texto(elementos[0], "CC_ZONA"),
                            prompt="Zona (opcional; en blanco = nivel del elemento):", title="Control de obra") or ""
sin_parametro = 0
with revit.Transaction("Control de obra: asignar partida"):
    for el in elementos:
        if not co.escribir(el, "CC_COD_PARTIDA", partida.strip()):
            sin_parametro += 1
            continue
        if actividad.strip():
            co.escribir(el, "CC_ID_ACTIVIDAD", actividad.strip())
        if zona.strip():
            co.escribir(el, "CC_ZONA", zona.strip())
        if not co.texto(el, "CC_ESTADO"):
            co.escribir(el, "CC_ESTADO", "Pendiente")
msg = u"Partida {} asignada a {} elemento(s).".format(partida.strip(), len(elementos) - sin_parametro)
if sin_parametro:
    msg += u"\n{} elemento(s) no tienen los parámetros CC_*: añade su categoría con 'Preparar modelo'.".format(sin_parametro)
forms.alert(msg, title="Control de obra")
