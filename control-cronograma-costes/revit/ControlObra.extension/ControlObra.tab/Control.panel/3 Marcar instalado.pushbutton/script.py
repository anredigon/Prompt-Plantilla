# -*- coding: utf-8 -*-
"""Marca los elementos seleccionados como Instalado (con la semana de obra) o los devuelve a Pendiente."""
from pyrevit import revit, forms
import control_obra as co

doc = revit.doc
elementos = revit.get_selection().elements
if not elementos:
    forms.alert("Selecciona primero los elementos montados.", exitscript=True)
estado = forms.CommandSwitchWindow.show(["Instalado", "Pendiente"], message=u"{} elemento(s) seleccionados. Marcar como:".format(len(elementos)))
if not estado:
    forms.alert("No se ha cambiado nada.", exitscript=True)
semana = None
if estado == "Instalado":
    inicio = co.inicio_obra(doc, preguntar=True)
    if inicio is None:
        forms.alert("Hace falta la fecha de inicio de obra para calcular la semana.", exitscript=True)
    propuesta = max(1, co.semana_actual(inicio))
    valor = forms.ask_for_string(default=str(propuesta), title="Control de obra",
                                 prompt=u"Semana de obra en que se instalaron (hoy es la semana {}; la semana 1 empezó el {}):"
                                 .format(propuesta, inicio.strftime("%d/%m/%Y")))
    try:
        semana = int(valor)
        assert 1 <= semana <= 104
    except Exception:
        forms.alert("La semana debe ser un número entre 1 y 104.", exitscript=True)
sin_partida, sin_parametro = 0, 0
with revit.Transaction(u"Control de obra: marcar {}".format(estado.lower())):
    for el in elementos:
        if not co.escribir(el, "CC_ESTADO", estado):
            sin_parametro += 1
            continue
        if semana is not None:
            co.escribir(el, "CC_SEMANA_INSTALADO", semana)
        if not co.texto(el, "CC_COD_PARTIDA"):
            sin_partida += 1
msg = u"{} elemento(s) marcados como {}{}.".format(len(elementos) - sin_parametro, estado,
                                                  u" en la semana {}".format(semana) if semana else u"")
if sin_partida:
    msg += u"\nAtención: {} no tienen partida asignada y la app no los contará. Usa 'Asignar partida'.".format(sin_partida)
if sin_parametro:
    msg += u"\n{} elemento(s) no tienen los parámetros CC_*: añade su categoría con 'Preparar modelo'.".format(sin_parametro)
forms.alert(msg, title="Control de obra")
