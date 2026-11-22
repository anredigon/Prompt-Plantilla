# Medidas DAX

Coinciden con el Excel (`Resumen`) y con el HTML. La semana de corte se lee de `parametros`.

```DAX
Semana corte = MAX ( parametros[semana_corte] )

BAC = SUM ( dim_partida[bac] )

-- Acumulados hasta la semana del contexto (para la curva S) ---------------
PV acum =
VAR s = MAX ( dim_semana[semana] )
RETURN CALCULATE ( SUM ( fact_semanal[pv_semana] ), dim_semana[semana] <= s, REMOVEFILTERS ( dim_semana ) )

EV acum =
VAR s = MAX ( dim_semana[semana] )
RETURN IF ( s <= [Semana corte],
    CALCULATE ( SUM ( fact_semanal[ev_semana] ), dim_semana[semana] <= s, REMOVEFILTERS ( dim_semana ) ) )

AC acum =
VAR s = MAX ( dim_semana[semana] )
RETURN IF ( s <= [Semana corte],
    CALCULATE ( SUM ( fact_planillas[importe] ), dim_semana[semana] <= s, REMOVEFILTERS ( dim_semana ) ) )

-- A fecha de corte (tarjetas y matriz) ----------------------------------
-- La semana de corte se guarda en una variable: un filtro booleano de CALCULATE no admite medidas
PV =
VAR c = [Semana corte]
RETURN CALCULATE ( SUM ( fact_semanal[pv_semana] ), dim_semana[semana] <= c, REMOVEFILTERS ( dim_semana ) )

EV =
VAR c = [Semana corte]
RETURN CALCULATE ( SUM ( fact_semanal[ev_semana] ), dim_semana[semana] <= c, REMOVEFILTERS ( dim_semana ) )

AC =
VAR c = [Semana corte]
RETURN CALCULATE ( SUM ( fact_planillas[importe] ), dim_semana[semana] <= c, REMOVEFILTERS ( dim_semana ) )

SV = [EV] - [PV]
CV = [EV] - [AC]
EV − AC = [EV] - [AC]
SPI = DIVIDE ( [EV], [PV] )
CPI = DIVIDE ( [EV], [AC] )
-- Sin CPI medible (EV = 0): coste real + trabajo pendiente a precio de oferta (igual que Excel y Python)
EAC =
VAR cpi = [CPI]
RETURN IF ( NOT ISBLANK ( cpi ) && cpi <> 0, DIVIDE ( [BAC], cpi ), [AC] + [BAC] - [EV] )
VAC = [BAC] - [EAC]
-- En blanco si ya está todo ganado o si se ha certificado toda la oferta quedando obra (inalcanzable)
TCPI =
VAR pendiente = [BAC] - [EV]
VAR margen = [BAC] - [AC]
RETURN IF ( pendiente > 0.01 && margen > 0, DIVIDE ( pendiente, margen ) )

TCPI inalcanzable = IF ( [BAC] - [EV] > 0.01 && [AC] >= [BAC], "Inalcanzable" )
% plan = DIVIDE ( [PV], [BAC] )
% real = DIVIDE ( [EV], [BAC] )

-- Earned schedule (convención de Lipke, igual que Python): C = última semana con PV acumulado <= EV,
-- interpolando con la siguiente ------------------------------------------------
ES =
VAR ev = CALCULATE ( [EV], ALL ( dim_partida ), ALL ( dim_actividad ) )
VAR t =
    MAXX (
        FILTER ( ALL ( dim_semana ),
            CALCULATE ( SUM ( fact_semanal[pv_semana] ), dim_semana[semana] <= EARLIER ( dim_semana[semana] ),
                        ALL ( dim_semana ), ALL ( dim_partida ), ALL ( dim_actividad ) ) <= ev ),
        dim_semana[semana] )
VAR t0 = COALESCE ( t, 0 )
VAR pv0 = CALCULATE ( SUM ( fact_semanal[pv_semana] ), dim_semana[semana] <= t0, ALL ( dim_semana ), ALL ( dim_partida ), ALL ( dim_actividad ) )
VAR pv1 = CALCULATE ( SUM ( fact_semanal[pv_semana] ), dim_semana[semana] <= t0 + 1, ALL ( dim_semana ), ALL ( dim_partida ), ALL ( dim_actividad ) )
RETURN t0 + DIVIDE ( ev - pv0, pv1 - pv0 )

-- Primera semana en que lo ganado alcanza el BAC (obra terminada); en blanco si no ha terminado
Semana fin real =
VAR bac = CALCULATE ( [BAC], ALL ( dim_partida ), ALL ( dim_actividad ) )
RETURN
    MINX (
        FILTER ( ALL ( dim_semana ),
            dim_semana[semana] <= [Semana corte]
                && CALCULATE ( SUM ( fact_semanal[ev_semana] ), dim_semana[semana] <= EARLIER ( dim_semana[semana] ),
                               ALL ( dim_semana ), ALL ( dim_partida ), ALL ( dim_actividad ) ) >= bac - 0.01 ),
        dim_semana[semana] )

SPI(t) =
VAR fin = [Semana fin real]
RETURN IF ( NOT ISBLANK ( fin ), DIVIDE ( MAX ( parametros[n_semanas] ), fin ),
            IF ( [PV] > 0, DIVIDE ( [ES], [Semana corte] ) ) )

Duración prevista (sem) =
VAR fin = [Semana fin real]
RETURN IF ( NOT ISBLANK ( fin ), fin, DIVIDE ( MAX ( parametros[n_semanas] ), [SPI(t)] ) )

Fin previsto =
VAR d = [Duración prevista (sem)]
RETURN IF ( NOT ISBLANK ( d ), MAX ( parametros[fecha_inicio] ) + ROUND ( 7 * d, 0 ) - 3 )

-- Pagos (lo certificado es lo que se debe; lo pagado se controla aparte) -----
Pagado =
VAR c = [Semana corte]
RETURN CALCULATE ( SUM ( fact_planillas[importe] ),
    FILTER ( ALL ( fact_planillas ), NOT ISBLANK ( fact_planillas[semana_pago] ) && fact_planillas[semana_pago] <= c ) )
Pendiente de pago = [AC] - [Pagado]

-- Fuera de oferta -----------------------------------------------------------
Extras sin decidir = CALCULATE ( SUM ( extras[importe] ), extras[estado] = "Sin decidir" )
Exposición máxima = [EAC] + [Extras sin decidir]
-- Con contingencia_pct = 0 (sin confirmar) el resultado es negativo si hay cambios propuestos: es intencionado
Contingencia disponible =
    MAX ( parametros[contingencia_pct] ) * CALCULATE ( [BAC], ALL ( dim_partida ) )
    - CALCULATE ( SUM ( contingencia[importe] ), contingencia[estado] IN { "Aprobado", "Propuesto" } )

-- Semáforo (mismos umbrales que el HTML: 0,95 proyecto, 0,90 actividades) ---
Nivel SPI proyecto =
VAR u = MAX ( parametros[umbral] )
VAR v = CALCULATE ( [SPI], ALL ( dim_actividad ), ALL ( dim_partida ) )
RETURN SWITCH ( TRUE (), ISBLANK ( v ), 0, v >= u, 4, v >= 0.85, 3, v >= 0.70, 2, 1 )

Nivel SPI =
VAR u = MAX ( parametros[umbral_actividad] )
VAR v = [SPI]
RETURN SWITCH ( TRUE (), ISBLANK ( v ), 0, v >= u, 4, v >= 0.85, 3, v >= 0.70, 2, 1 )
-- 4 OK · 3 Vigilar · 2 Desviado · 1 Crítico · 0 No iniciada
```
