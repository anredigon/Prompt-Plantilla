# Power BI · montaje de la maqueta

Aquí no se puede generar un `.pbix` (Power BI Desktop solo corre en Windows). Lo que sí está listo: el modelo de datos en CSV, las medidas DAX y el tema de colores. Montarlo lleva entre 1 y 2 horas.

## 1. Cargar datos
`Obtener datos → Texto/CSV` para cada archivo de esta carpeta. Al importar, en **Configuración regional del archivo** elige *Inglés (Estados Unidos)*: los decimales vienen con punto.

| Tabla | Grano | Clave |
|---|---|---|
| `dim_partida` | partida de oferta | `cod_partida` |
| `dim_actividad` | actividad del cronograma | `id_actividad` |
| `dim_semana` | semana del proyecto | `semana` |
| `fact_semanal` | partida × semana (PV, EV y AC incrementales y acumulados) | `cod_partida`, `semana` |
| `fact_planillas` | línea de planilla | `cod_partida`, `semana` |
| `extras`, `contingencia`, `parametros` | tablas sueltas | — |

## 2. Relaciones (todas 1:* con filtro en una dirección)
- `dim_partida[cod_partida]` → `fact_semanal[cod_partida]` y → `fact_planillas[cod_partida]`
- `dim_actividad[id_actividad]` → `dim_partida[id_actividad]`
- `dim_semana[semana]` → `fact_semanal[semana]` y → `fact_planillas[semana]`

## 3. Medidas
Copia las de [`MEDIDAS_DAX.md`](MEDIDAS_DAX.md) en una tabla vacía `_Medidas`.

## 4. Tema
`Vista → Temas → Buscar temas → tema_powerbi.json`. Usa la misma paleta que el HTML, validada para daltonismo. Asigna los colores a mano en la curva S: PV en azul (#2a78d6), AC en naranja (#eb6834) y EV en verde agua (#1baf7a).

## 5. Página única para el directorio (misma estructura que el HTML)
1. **Tarjetas**: Fin previsto · EAC · Exposición máxima. Debajo: SPI, CPI, Certificado − Ganado (con Pagado y Pendiente de pago), TCPI y Contingencia disponible.
2. **Gráfico de líneas** con eje `dim_semana[semana]` y valores `[PV acum]`, `[EV acum]` y `[AC acum]`. Añade una línea constante con `[BAC]`.
3. **Matriz** con filas `dim_actividad[actividad]` y columnas `[% plan]`, `[% real]`, `[SPI]`, `[CPI]`, `[EV − AC]`. Aplica formato condicional de iconos a la columna SPI usando `[Nivel SPI]` (umbral 0,90). Para las tarjetas del proyecto usa `[Nivel SPI proyecto]` (umbral 0,95).
4. **Tabla** `extras` filtrada a `estado = "Sin decidir"`.
5. **Segmentación** por `dim_partida[capitulo]`.

## 6. Actualización semanal
Actualiza el Excel y ejecuta `python scripts/generar_dashboard.py`, que regenera los CSV. Luego pulsa **Actualizar** en Power BI.

Siguiente paso, cuando deje de ser maqueta: que Power Query lea directamente el Excel (o la exportación de Revit en SharePoint) y eliminar el paso intermedio del script.
