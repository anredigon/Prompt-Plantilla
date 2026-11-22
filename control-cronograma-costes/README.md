# Control semanal de cronograma y costes desde Revit (maqueta)

Informe semanal al directorio con valor ganado (EVM): avance físico medido en unidades del modelo Revit, coste real desde planillas y línea base = oferta `Costes GPT 09 0008.pdf` (1.375.000 €).

> **MAQUETA.** Los importes de la oferta son reales (tomados del PDF). Cronograma, exportación Revit, avance y planillas son **ficticios**. Sirven para validar formato y método antes de cargar datos reales.

## Contenido
| Archivo | Para qué |
|---|---|
| `datos/Control_Proyecto_Maqueta.xlsx` | **Fuente única**. Oferta, cronograma, exportación Revit, avance manual, planillas, extras y contingencia. Se calcula solo con fórmulas (hoja `Resumen`). |
| `dashboard_directorio.html` | Informe del directorio. Un solo archivo, se abre en el navegador, se imprime en PDF. |
| `powerbi/` | CSV en estrella, medidas DAX, tema y guía de montaje. |
| `scripts/crear_maqueta_excel.py` | Regenera el Excel de la maqueta con datos ficticios. |
| `scripts/generar_dashboard.py` | Lee el Excel y genera el HTML y los CSV de Power BI. |

## Ciclo semanal
1. En Revit, actualizar `CC_ESTADO` y `CC_SEMANA_INSTALADO` de lo montado. Exportar la tabla de planificación y pegarla en la hoja `Revit_Export`.
2. Hoja `Avance_Manual`: % acumulado de lo que no está modelado (protecciones, cargas, puesta en marcha).
3. Hoja `Planillas`: líneas certificadas en la semana.
4. Hoja `Parametros`: subir la semana de corte.
5. `python scripts/generar_dashboard.py` y enviar el HTML. En Power BI, pulsar Actualizar.

## Parámetros compartidos a crear en Revit (hoy no existen)
Sin esto no hay enlace modelo ↔ coste ↔ plazo. Crear como **parámetros compartidos de ejemplar** en las categorías Mechanical Equipment, Pipes, Pipe Insulations, Generic Models y Electrical Equipment (y en el resto de disciplinas cuando entren):

| Parámetro | Tipo | Ejemplo | Uso |
|---|---|---|---|
| `CC_COD_PARTIDA` | Texto | `P11` | Enlace con la partida de oferta/presupuesto |
| `CC_ID_ACTIVIDAD` | Texto | `ACT-070` | Enlace con la actividad del cronograma |
| `CC_ESTADO` | Texto (lista: Pendiente / Instalado) | `Instalado` | Avance físico |
| `CC_SEMANA_INSTALADO` | Entero | `9` | Curva S semanal |
| `CC_ZONA` | Texto (opcional) | `Muelles` | Filtros por zona |

Tabla de planificación a exportar, con columnas en este orden: `ElementId, Categoría, Familia:Tipo, Zona, CC_COD_PARTIDA, CC_ID_ACTIVIDAD, Cantidad (1 o Longitud), Ud, CC_ESTADO, CC_SEMANA_INSTALADO`.

Control de calidad que aplica la maqueta: la oferta tiene **52 desagües y 52 soportes, que corresponden a 52 evaporadores** (12 de cámaras, 30 de muelles, 2 de despojos y 8 de frescos). El modelo debe contar exactamente esos elementos. Si no salen, o el modelo o la oferta están mal.

## Reglas de cálculo
- **PV**: BAC de cada partida repartido linealmente entre el inicio y el fin de su actividad.
- **EV**: unidades instaladas ÷ unidades del modelo × BAC. Partidas fuera del modelo: % manual. Montaje (global): proporcional al nº de elementos del modelo de frío ya montados. No se calcula por €, porque los compresores lo inflarían.
- **AC**: importe certificado en planillas. **AC > EV significa que se paga por delante de lo instalado** (acopios, sobremediciones, precios contradictorios).
- **SPI(t)** por *earned schedule*. Fin previsto = duración ÷ SPI(t).
- **EAC** = BAC ÷ CPI. **Exposición máxima** = EAC + extras sin decidir.
- Semáforo: ≥ 0,95 OK · ≥ 0,85 Vigilar · ≥ 0,70 Desviado · < 0,70 Crítico. Un EAC por encima de la oferta nunca se muestra en verde.

## Supuestos a confirmar
- Contingencia del 5 % (no figura en la oferta).
- Cantidades de medición de tuberías y aislamiento (1.850 m) y del cuadro eléctrico (6 módulos). La oferta las da como 1 ud global.
- Semanas límite de decisión de los extras.
