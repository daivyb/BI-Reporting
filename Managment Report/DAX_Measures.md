# Advanced DAX Measures for Kimball Star Schema

This document highlights the complex DAX measures developed for this project. The core challenge was that the physical progress data (Sciforma) was tracked strictly at the lowest Technical level (WBS Task), while the financial budget data (EMMA) was tracked at various administrative and technical levels. 

By restructuring the data into a Kimball Star Schema with a unified Technical Dimension (`Dim_Tecnica`), we used DAX to dynamically calculate the physical progress (Avance Físico) and planned goals (Meta Física) across the entire WBS hierarchy (Project > Component > Subcomponent > Activity > Task) using nested iterators.

## 1. Physical Progress (Avance Físico Base)

This measure uses `ISINSCOPE` to detect the current level of the matrix visual and recursively averages the physical progress from the task level up to the project level.

```dax
Avance Fisico Base = 
SWITCH(
    TRUE(),
    -- 1. Task Level: Direct percentage
    ISINSCOPE('Dim_Tecnica'[Tarea]), 
        AVERAGE('Fact_Avance_Fisico'[Pct_Avance_Fisico]),
        
    -- 2. Activity Level: Average of its Tasks
    ISINSCOPE('Dim_Tecnica'[Actividad]), 
        AVERAGEX(
            VALUES('Dim_Tecnica'[Tarea]), 
            CALCULATE(AVERAGE('Fact_Avance_Fisico'[Pct_Avance_Fisico]))
        ),
        
    -- 3. Subcomponent Level: Average of its Activities
    ISINSCOPE('Dim_Tecnica'[Subcomponente]), 
        AVERAGEX(
            VALUES('Dim_Tecnica'[Actividad]),
            CALCULATE(
                AVERAGEX(
                    VALUES('Dim_Tecnica'[Tarea]), 
                    CALCULATE(AVERAGE('Fact_Avance_Fisico'[Pct_Avance_Fisico]))
                )
            )
        ),
        
    -- 4. Component Level: Average of its Subcomponents
    ISINSCOPE('Dim_Tecnica'[Componente]), 
        AVERAGEX(
            VALUES('Dim_Tecnica'[Subcomponente]),
            CALCULATE(
                AVERAGEX(
                    VALUES('Dim_Tecnica'[Actividad]),
                    CALCULATE(
                        AVERAGEX(
                            VALUES('Dim_Tecnica'[Tarea]), 
                            CALCULATE(AVERAGE('Fact_Avance_Fisico'[Pct_Avance_Fisico]))
                        )
                    )
                )
            )
        ),
        
    -- 5. Project Level: Average of its Components
    ISINSCOPE('Dim_Tecnica'[Proyecto]), 
        AVERAGEX(
            VALUES('Dim_Tecnica'[Componente]),
            CALCULATE(
                AVERAGEX(
                    VALUES('Dim_Tecnica'[Subcomponente]),
                    CALCULATE(
                        AVERAGEX(
                            VALUES('Dim_Tecnica'[Actividad]),
                            CALCULATE(
                                AVERAGEX(
                                    VALUES('Dim_Tecnica'[Tarea]), 
                                    CALCULATE(AVERAGE('Fact_Avance_Fisico'[Pct_Avance_Fisico]))
                                )
                            )
                        )
                    )
                )
            )
        ),
        
    BLANK()
)
```

## 2. Planned Physical Goal (Meta Física Base)

The planned goal requires the classic DAX pattern of "Summing over time, but Averaging over the hierarchy". The base metric is `SUM` at the task level so it accumulates across selected quarters, but it uses `AVERAGEX` as it moves up the WBS hierarchy.

```dax
Meta Fisica Base = 
SWITCH(
    TRUE(),
    -- 1. Task Level: SUM over time
    ISINSCOPE('Dim_Tecnica'[Tarea]), 
        SUM('Fact_Avance_Fisico'[Meta_Fisica]),
        
    -- 2. Activity Level: Average of the summed Tasks
    ISINSCOPE('Dim_Tecnica'[Actividad]), 
        AVERAGEX(
            VALUES('Dim_Tecnica'[Tarea]), 
            CALCULATE(SUM('Fact_Avance_Fisico'[Meta_Fisica]))
        ),
        
    -- 3. Subcomponent Level: Average of its Activities
    ISINSCOPE('Dim_Tecnica'[Subcomponente]), 
        AVERAGEX(
            VALUES('Dim_Tecnica'[Actividad]),
            CALCULATE(
                AVERAGEX(
                    VALUES('Dim_Tecnica'[Tarea]), 
                    CALCULATE(SUM('Fact_Avance_Fisico'[Meta_Fisica]))
                )
            )
        ),
        
    -- 4. Component Level: Average of its Subcomponents
    ISINSCOPE('Dim_Tecnica'[Componente]), 
        AVERAGEX(
            VALUES('Dim_Tecnica'[Subcomponente]),
            CALCULATE(
                AVERAGEX(
                    VALUES('Dim_Tecnica'[Actividad]),
                    CALCULATE(
                        AVERAGEX(
                            VALUES('Dim_Tecnica'[Tarea]), 
                            CALCULATE(SUM('Fact_Avance_Fisico'[Meta_Fisica]))
                        )
                    )
                )
            )
        ),
        
    -- 5. Project Level: Average of its Components
    ISINSCOPE('Dim_Tecnica'[Proyecto]), 
        AVERAGEX(
            VALUES('Dim_Tecnica'[Componente]),
            CALCULATE(
                AVERAGEX(
                    VALUES('Dim_Tecnica'[Subcomponente]),
                    CALCULATE(
                        AVERAGEX(
                            VALUES('Dim_Tecnica'[Actividad]),
                            CALCULATE(
                                AVERAGEX(
                                    VALUES('Dim_Tecnica'[Tarea]), 
                                    CALCULATE(SUM('Fact_Avance_Fisico'[Meta_Fisica]))
                                )
                            )
                        )
                    )
                )
            )
        ),
        
    BLANK()
)
```

## 3. Physical Gap (Brecha Física)

To prevent Power BI's "Dense Evaluation" (cross-joining) from rendering non-existent combinations of administrative dimensions and technical components, the gap mathematical calculation is wrapped in a blank check `IF(NOT ISBLANK(...))`. This ensures that "Phantom" projects or components are dynamically hidden from visuals.

```dax
Brecha Fisica (%) = 
VAR AvanceAcumulado = [Avance Fisico Base]
VAR MetaAcumulada = TOTALYTD([Meta Fisica Base], 'Calendario'[Date])

RETURN
    -- Protection against DAX Cartesian products evaluating to 100% gap
    IF(
        NOT ISBLANK(MetaAcumulada), 
        1 - DIVIDE(AvanceAcumulado, MetaAcumulada, 0)
    )
```

## 4. Dynamic Tooltip Title

Used to provide a contextual title to tooltip line charts (S-Curves), updating dynamically based on the project hovered over in the main visual.

```dax
Titulo_Dinamico_Tooltip = 
VAR NombreProyecto = SELECTEDVALUE('Dim_Administrativa'[Proyecto], "Portafolio Total")

RETURN 
    "Desempeño Físico Acumulado (Avance vs. Meta) | " & NombreProyecto
```
