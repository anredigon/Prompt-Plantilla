# Control semanal de cronograma y costes (valor ganado) · plantilla reutilizable

Herramienta para controlar cualquier proyecto con valor ganado (EVM) e informar cada semana al directorio. El avance físico sale de Revit y de los albaranes, y el coste real de las planillas.

**Cómo funciona:** cada proyecto tiene su propio Excel en su carpeta de Drive (los datos). Una única app HTML lo lee y genera el informe (la vista). No hace falta instalar nada ni hay servidor, y los datos no salen de tu ordenador.

## Contenido
| Archivo | Para qué |
|---|---|
| `App_Control_Proyecto.html` | **La app.** Ábrela con doble clic, carga el Excel del proyecto y obtienes el informe. Funciona sin conexión. |
| `plantilla/Plantilla_Control_Proyecto.xlsx` | **Plantilla en blanco.** Haz una copia por proyecto. Admite hasta 300 partidas, 150 actividades, 5.000 elementos de Revit y 104 semanas. |
| `ejemplo/Ejemplo_Control_Proyecto.xlsx` | Ejemplo relleno (costes de la oferta «Costes GPT 09 0008»; avance, Revit y planillas **ficticios**). |
| `ejemplo/Informe_Ejemplo_S11.html` | Informe exportado desde la app con el ejemplo. Así es lo que recibe el directorio. |
| `ejemplo/PowerBI_Ejemplo.xlsx` | Datos para Power BI exportados desde la app con el ejemplo. |
| `powerbi/` | Guía de montaje, medidas DAX y tema de colores. |
| `scripts/` | Mantenimiento: generar la plantilla (`crear_plantilla.py`), ensamblar la app (`construir_app.py`) y probarla (`probar_app.js`). No hacen falta para usarla. |

## Empezar un proyecto nuevo
1. Copia `plantilla/Plantilla_Control_Proyecto.xlsx` en la carpeta de Drive del proyecto y renómbrala `Control_<Proyecto>.xlsx`.
2. Rellena **solo las celdas en azul**:
   - `Parametros`: nombre, lunes de la semana 1, nº de semanas, contingencia y umbrales.
   - `Cronograma`: actividades (ID, nombre, semana de inicio y fin).
   - `Oferta`: partidas (código, capítulo, importe, actividad, fuente de avance, unidad y cantidad, perfil del planificado).
   - `Horas_Montaje`: horas por unidad de lo que se monta (si hay partida de montaje `DERIVADO`).
3. En Revit, crea los parámetros `CC_*` (abajo).
4. Abre `App_Control_Proyecto.html` y carga el Excel. Si algo está mal enlazado, la app lo avisa arriba del informe.

## Ciclo semanal
1. Revit: actualiza `CC_ESTADO` y `CC_SEMANA_INSTALADO`, exporta la tabla de planificación y pégala en `Revit_Export` (columnas A a J).
2. `Suministros`: albaranes de entrega de la semana.
3. `Avance_Manual`: % acumulado de lo que no está en el modelo.
4. `Planillas`: líneas certificadas y, cuando se paguen, su **semana de pago**.
5. `Parametros`: sube la semana de corte y guarda.
6. App: carga el Excel y usa:
   - **Exportar informe**: un HTML ligero para enviar al directorio, que se abre en cualquier navegador.
   - **Imprimir / PDF**.
   - **Datos Power BI**.

   El selector **Semana de corte** permite ver semanas anteriores sin tocar el Excel.

## Permisos para el equipo
Se gestionan en Drive, compartiendo la carpeta del proyecto:
- **Lector**: ve el Excel y los informes exportados, y puede abrirlos en la app. Es lo recomendado para compañeros y dirección.
- **Editor**: puede modificar el Excel. Que edite **una sola persona a la vez**: Drive para escritorio no fusiona cambios de Excel y crearía copias en conflicto.

La app no guarda nada ni envía datos a ningún sitio, así que no necesita permisos propios.

## Parámetros compartidos en Revit
Créalos como **parámetros compartidos de ejemplar** en las categorías que se controlen:

| Parámetro | Tipo | Ejemplo | Uso |
|---|---|---|---|
| `CC_COD_PARTIDA` | Texto | `P11` | Enlace con la partida de la oferta |
| `CC_ID_ACTIVIDAD` | Texto | `ACT-070` | Enlace con la actividad del cronograma |
| `CC_ESTADO` | Texto (Pendiente / Instalado) | `Instalado` | Avance físico |
| `CC_SEMANA_INSTALADO` | Entero | `9` | Curva S semanal |
| `CC_ZONA` | Texto (opcional) | `Muelles` | Filtros por zona |

La tabla de planificación se exporta con estas columnas, en este orden: `ElementId, Categoría, Familia:Tipo, Zona, CC_COD_PARTIDA, CC_ID_ACTIVIDAD, Cantidad (1 o Longitud), Ud, CC_ESTADO, CC_SEMANA_INSTALADO`.

## Reglas de cálculo (validadas con dirección)
- **PV**: perfil `ENTREGA` (equipos) al 100 % en su semana de entrega prevista (columna X de `Oferta`; en blanco = inicio de la actividad). Perfil `LINEAL` repartido por igual en la duración de su actividad.
- **EV**:
  - `ALBARAN`: entregado según la hoja `Suministros`.
  - `REVIT`: instalado según el modelo.
  - `MANUAL`: % semanal.
  - `DERIVADO` (montaje): horas de los elementos instalados ÷ horas totales.
  - Tope del 100 % por partida: lo que se entregue o instale de más se alerta como alcance.
- **AC**: lo certificado en planillas, es decir, lo que se debe. Lo **pagado** y lo **pendiente de pago** se muestran aparte. Si AC > EV, se está certificando por delante de lo hecho.
- **SPI(t)** por *earned schedule*: el momento más temprano en que el plan alcanzaba lo ganado hoy (las semanas sin nada planificado no cuentan como avance). Fin previsto = duración ÷ SPI(t). Con la obra terminada se muestra el fin real. Si la obra sigue después del fin planificado, la semana de corte puede superar el nº de semanas (hasta 104).
- **EAC** = BAC ÷ CPI; sin CPI medible, AC + (BAC − EV). **Exposición máxima** = EAC + extras sin decidir.
- **Semáforo**: umbral 0,95 para el proyecto y 0,90 por actividad. Por debajo: ≥ 0,85 Vigilar, ≥ 0,70 Desviado, < 0,70 Crítico. Un EAC por encima de la oferta nunca sale en verde.
- **Contingencia**: 0 % hasta que dirección confirme que existe y su importe.
- **Comprobaciones al cargar** (avisos arriba del informe): códigos de partida o actividad repetidos, partidas sin actividad, estados de Revit no reconocidos, elementos instalados sin semana, % manual fuera de 0–100 %, números guardados como texto y pagos anteriores a su certificación.

## Supuestos del ejemplo que hay que confirmar en un proyecto real
- Horas de montaje por tipo de elemento: pedirlas al instalador.
- Semanas de entrega pactadas con los proveedores y plazo de pago.
- Existencia e importe de la contingencia.

## Mantenimiento (solo si cambias la herramienta)
```
python scripts/crear_plantilla.py              # plantilla en blanco
python scripts/crear_plantilla.py --ejemplo    # ejemplo relleno
python scripts/construir_app.py                # ensambla App_Control_Proyecto.html
node scripts/probar_app.js                     # prueba la app con el ejemplo (requiere Playwright)
```
La app usa SheetJS (`scripts/vendor/xlsx.full.min.js`, licencia Apache-2.0) para leer Excel en el navegador.
