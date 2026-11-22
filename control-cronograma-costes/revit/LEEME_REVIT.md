# Botones de Revit · Control de obra (pyRevit)

Cuatro botones en una pestaña **ControlObra** de Revit. Conectan el modelo con la app de control de obra.

| Botón | Cuándo | Qué hace |
|---|---|---|
| **1 · Preparar modelo** | Una vez por modelo | Crea los parámetros `CC_COD_PARTIDA`, `CC_ID_ACTIVIDAD`, `CC_ESTADO`, `CC_SEMANA_INSTALADO` y `CC_ZONA` en las categorías que elijas. En Información de proyecto crea `CC_INICIO_OBRA` (lunes de la semana 1) y `CC_CARPETA_EXPORT`. |
| **2 · Asignar partida** | Al empezar, y cuando se modele algo nuevo | Seleccionas elementos y escribes su código de partida del presupuesto (y opcionalmente actividad y zona). |
| **3 · Marcar instalado** | Cada semana | Seleccionas lo montado y lo marcas **Instalado** con la semana de obra, que propone sola, o lo devuelves a **Pendiente**. |
| **4 · Exportar avance** | Cada semana | Genera `Revit_Avance_<modelo>.csv` en la carpeta del proyecto. En la app: **Actualizar semana → Avance en Revit →** adjunta ese archivo. |

## Instalación (una vez, 10 minutos)
1. Instala **pyRevit** (gratuito) desde https://pyrevitlabs.io, versión estable, con Revit cerrado.
2. Copia la carpeta `revit` (la que contiene `ControlObra.extension`) a un sitio fijo. Por ejemplo, la carpeta de la herramienta en Drive: `…\Dashboard_Control_Cronograma_y_Costos\revit`.
3. Abre Revit → pestaña **pyRevit** → **Settings** → **Custom Extension Directories** → **Add folder**. Elige esa carpeta `revit` → **Save Settings and Reload**.
4. Aparece la pestaña **ControlObra** con los cuatro botones.

## Primer uso en un modelo
1. **Preparar modelo**: marca las categorías que vas a controlar (equipos mecánicos, tuberías, aislamiento…) e indica la fecha de inicio de obra.
2. Selecciona elementos, por ejemplo con *clic derecho → Seleccionar todos los ejemplares*, y usa **Asignar partida** con el código del presupuesto que tiene ese proyecto en la app (P01, P11…).
3. Guarda el modelo.

## Cada semana (2 minutos)
1. Selecciona lo que se ha montado → **Marcar instalado** → acepta la semana propuesta.
2. **Exportar avance**. La primera vez pide la carpeta del proyecto y la recuerda.
3. En la app: **Actualizar semana** → adjunta `Revit_Avance_<modelo>.csv` → **Guardar semana**.

## Cómo mide
- Tuberías, conductos, aislamiento, bandejas y tubos se miden en **metros**, con la longitud del elemento. El resto, en **unidades**.
- Si en la app una partida se mide en metros, su "cantidad de medición" debe ser los metros totales del modelo.

## Limitaciones conocidas
- **No se ha podido probar dentro de Revit en este entorno.** El código está escrito para pyRevit (IronPython 2.7 y CPython 3) y Revit 2020–2025, pero la primera vez pruébalo en una copia del modelo. Si un botón falla, copia el mensaje de error y envíamelo.
- Los elementos de **vínculos de Revit** (modelos vinculados) no se exportan: los botones trabajan sobre el modelo abierto.
- Los elementos de grupos y piezas se exportan, pero comprueba que tienen la partida asignada.
