# Control de obra

Una app para controlar plazo y coste de cada obra e informar al directorio cada semana.

## Uso
1. Doble clic en **`App_Control_Obra.html`** (con Google Chrome o Microsoft Edge).
2. La primera vez, **elige tu carpeta de proyectos** (por ejemplo la de Drive). La app la recuerda.
3. **Elige un proyecto** de la lista o crea uno nuevo. Cada proyecto es una subcarpeta con:
   - `Control_<proyecto>.xlsx`: todos los datos (se puede abrir en Excel).
   - `Documentos/`: albaranes, certificaciones y documentos que subes.
   - `Informes/`: los informes semanales que guardas.
4. Pestañas del proyecto:
   - **Informe**: plazo, coste final, decisiones y desviaciones. *Guardar informe en la carpeta* o *Imprimir / PDF*.
   - **Actualizar semana**: subes el CSV de Revit, los albaranes y la certificación (se guardan en la carpeta y tecleas sus cantidades), el % de lo que no está en Revit, los pagos y las decisiones sobre extras.
   - **Datos del proyecto**: datos generales, cronograma y presupuesto (se pueden importar desde Excel o CSV), extras y contingencia.
   - **Documentos**: subir y abrir los documentos del proyecto.

Para probarla: elige como carpeta de proyectos esta misma carpeta; aparecerá el proyecto **ejemplo** (datos de avance ficticios sobre la oferta «Costes GPT 09 0008»).

## Revit
Botones pyRevit para marcar lo instalado y exportar el avance que lee la app: `revit/LEEME_REVIT.md`.

## Cómo calcula
- Plan: equipos al 100 % en su semana de entrega; resto repartido en su actividad.
- Hecho: equipos por albarán; elementos del modelo por lo marcado como instalado en Revit; montaje por horas instaladas; resto por % semanal. Máximo 100 % por partida.
- Coste real = lo certificado; lo pagado y lo pendiente, aparte.
- Fin previsto = duración ÷ ritmo real (earned schedule). Coste final = presupuesto ÷ eficiencia de coste actual.
- Alertas: proyecto por debajo de 0,95; actividad por debajo de 0,90.

## Mantenimiento
`python scripts/construir_app.py` ensambla `App_Control_Obra.html` a partir de `scripts/app/` (motor de cálculo, informe, utilidades y la app) y SheetJS (`scripts/vendor/`, Apache-2.0). `python scripts/crear_plantilla.py --ejemplo` regenera el Excel de ejemplo.
