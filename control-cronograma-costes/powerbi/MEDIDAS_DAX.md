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
PV = CALCULATE ( SUM ( fact_semanal[pv_semana] ), dim_semana[semana] <= [Semana corte], REMOVEFILTERS ( dim_semana ) )
EV = CALCULATE ( SUM ( fact_semanal[ev_semana] ), dim_semana[semana] <= [Semana corte], REMOVEFILTERS ( dim_semana ) )
AC = CALCULATE ( SUM ( fact_planillas[importe] ), dim_semana[semana] <= [Semana corte], REMOVEFILTERS ( dim_semana ) )

SV = [EV] - [PV]
CV = [EV] - [AC]
EV − AC = [EV] - [AC]
SPI = DIVIDE ( [EV], [PV] )
CPI = DIVIDE ( [EV], [AC] )
EAC = DIVIDE ( [BAC], [CPI], [BAC] )
VAC = [BAC] - [EAC]
TCPI = DIVIDE ( [BAC] - [EV], [BAC] - [AC] )
% plan = DIVIDE ( [PV], [BAC] )
% real = DIVIDE ( [EV], [BAC] )

-- Earned schedule: semana en que lo ganado hoy debía estar ganado ----------
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

SPI(t) = DIVIDE ( [ES], [Semana corte] )
Duración prevista (sem) = DIVIDE ( MAX ( parametros[n_semanas] ), [SPI(t)] )
Fin previsto =
    MAX ( parametros[fecha_inicio] ) + ROUND ( 7 * [Duración prevista (sem)], 0 ) - 3

-- Fuera de oferta -----------------------------------------------------------
Extras sin decidir = CALCULATE ( SUM ( extras[importe] ), extras[estado] = "Sin decidir" )
Exposición máxima = [EAC] + [Extras sin decidir]
Contingencia disponible =
    MAX ( parametros[contingencia_pct] ) * CALCULATE ( [BAC], ALL ( dim_partida ) )
    - CALCULATE ( SUM ( contingencia[importe] ), contingencia[estado] IN { "Aprobado", "Propuesto" } )

-- Semáforo (mismo umbral que el HTML) ---------------------------------------
Nivel SPI =
VAR u = MAX ( parametros[umbral] )
VAR v = [SPI]
RETURN SWITCH ( TRUE (), ISBLANK ( v ), 0, v >= u, 4, v >= 0.85, 3, v >= 0.70, 2, 1 )
-- 4 OK · 3 Vigilar · 2 Desviado · 1 Crítico · 0 No iniciada
```
