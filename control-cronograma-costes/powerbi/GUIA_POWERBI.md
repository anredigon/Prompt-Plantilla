# Power BI · plantilla reutilizable

No puedo generar el `.pbix` (Power BI Desktop solo funciona en Windows). Están listos el modelo de datos, las medidas DAX y el tema de colores. El montaje se hace **una sola vez**, entre 1 y 2 horas. Después lo guardas como plantilla (`.pbit`) y lo reutilizas en cada proyecto.

## 0. De dónde salen los datos
En la app (`App_Control_Proyecto.html`), con el proyecto cargado, pulsa **Datos Power BI**. Se descarga `PowerBI_<Proyecto>.xlsx`, con una hoja por tabla. Guárdalo en la carpeta del proyecto con ese mismo nombre y sobrescríbelo cada semana: así en Power BI solo tienes que pulsar **Actualizar**.
Configura el navegador para que pregunte dónde guardar las descargas; si no, irá a la carpeta Descargas.

## 1. Cargar datos
`Obtener datos → Libro de Excel` → `PowerBI_<Proyecto>.xlsx` → marca las 8 hojas. Puedes probar con `ejemplo/PowerBI_Ejemplo.xlsx`.

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
Actualiza el Excel del proyecto, ábrelo en la app, pulsa **Datos Power BI**, sobrescribe `PowerBI_<Proyecto>.xlsx` y pulsa **Actualizar** en Power BI.

## 7. Reutilizar en otros proyectos
`Archivo → Exportar → Plantilla de Power BI (.pbit)`. Para un proyecto nuevo, abre el `.pbit` y, en `Transformar datos → Configuración de origen de datos`, apunta al `PowerBI_<Proyecto>.xlsx` de ese proyecto.

## Control de calidad
La primera vez, compara las tarjetas de Power BI con el informe de la app o con la hoja `Resumen` del Excel. Tienen que cuadrar al euro (en el ejemplo: EAC 1.422.506 €, SPI 0,88, CPI 0,97). Si no cuadran, no lo presentes.
