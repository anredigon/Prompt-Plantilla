# -*- coding: utf-8 -*-
"""Exporta Revit_Avance_<modelo>.csv con todos los elementos que tienen CC_COD_PARTIDA."""
import io
import os
import re

from pyrevit import revit, forms, DB
import control_obra as co

doc = revit.doc
info = doc.ProjectInformation
carpeta = co.texto(info, "CC_CARPETA_EXPORT")
if not carpeta or not os.path.isdir(carpeta):
    forms.alert("Elige la carpeta del proyecto (por ejemplo la de Google Drive) donde guardar el avance. Se recordará para la próxima vez.",
                title="Control de obra")
    carpeta = forms.pick_folder(title="Carpeta del proyecto para el archivo de avance")
    if not carpeta:
        forms.alert("No se ha exportado nada.", exitscript=True)
    with revit.Transaction("Control de obra: carpeta de exportación"):
        co.escribir(info, "CC_CARPETA_EXPORT", carpeta)

filas, instalados, sin_semana = [], 0, 0
for el in DB.FilteredElementCollector(doc).WhereElementIsNotElementType():
    partida = co.texto(el, "CC_COD_PARTIDA")
    if not partida:
        continue
    q, ud = co.cantidad(el)
    estado = co.texto(el, "CC_ESTADO") or u"Pendiente"
    semana = co.entero(el, "CC_SEMANA_INSTALADO")
    if estado.lower() == u"instalado":
        instalados += 1
        if not semana:
            sin_semana += 1
    filas.append([co.id_entero(el.Id), el.Category.Name if el.Category else u"", co.familia_tipo(doc, el), co.zona(doc, el),
                  partida, co.texto(el, "CC_ID_ACTIVIDAD"), q, ud, estado, semana if semana else u""])
if not filas:
    forms.alert("Ningún elemento tiene CC_COD_PARTIDA. Asigna partidas con 'Asignar partida' antes de exportar.", exitscript=True)

nombre = re.sub(r"[^\w\-]+", "_", doc.Title).strip("_") or "modelo"
ruta = os.path.join(carpeta, u"Revit_Avance_{}.csv".format(nombre))
with io.open(ruta, "w", encoding="utf-8-sig", newline="") as f:
    f.write(co.csv_linea(co.CABECERA_CSV))
    for fila in filas:
        f.write(co.csv_linea(fila))
msg = u"Exportados {} elementos ({} instalados) a:\n{}\n\nEn la app: Actualizar semana > Avance en Revit > adjunta este archivo.".format(len(filas), instalados, ruta)
if sin_semana:
    msg += u"\n\nAtención: {} elementos 'Instalado' no tienen semana y no sumarán avance. Vuelve a marcarlos con 'Marcar instalado'.".format(sin_semana)
forms.alert(msg, title="Control de obra")
