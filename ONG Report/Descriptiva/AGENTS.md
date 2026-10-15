# Profonanpe vTrack BI - Reglas del Proyecto

Este directorio contiene la arquitectura de datos y los procesos de transformación (ETL) para el dashboard de vTrack BI de Profonanpe en Power BI.

## Arquitectura del Modelo de Datos (Metodología Kimball)

Este proyecto emplea un modelo dimensional puro (Star Schema) diseñado para maximizar el rendimiento en el motor VertiPaq de Power BI y evitar ciclos o relaciones inactivas:

1.  **Tablas de Dimensiones (1:1)**:
    Las entidades de negocio (Proyectos, Subdonaciones, Ámbito) viven en dimensiones. Cualquier característica que tenga una relación 1 a 1 con el proyecto (ej. Rol de PROFONANPE, Área Temática, Tipo de Intervención, Estado Plan Operativo Anual) **debe desnormalizarse** como columna de texto dentro de la dimensión principal (`Dim_Proyectos` o `Dim_Subdonacion`), en lugar de crear patrones Snowflake innecesarios.
    *   **Jerarquía Proyectos ──> Subdonaciones**: `Dim_Subdonacion` conserva la llave `IdProyecto` para conectarse como relación 1 a Varios (1:*) directa con `Dim_Proyectos`.

2.  **Dimensiones de Ubicación Independientes (Macro vs Micro)**:
    *   `Dim_Ubicacion_Proyecto`: Contiene únicamente los 25 Departamentos a nivel macro vinculados a proyectos.
    *   `Dim_Ubicacion_Subdonacion`: Contiene los Distritos (con su Provincia y Departamento) a nivel micro para las subdonaciones.
    *   *Razón arquitectónica*: Mantenerlas separadas previene que Power BI detecte bucles cerrados y desactive las relaciones puente (Bridge) en los mapas.

3.  **Tablas de Hechos (Facts)**:
    Tablas como `Fact_Proyectos` o `Fact_Subdonacion` contienen **exclusivamente** llaves foráneas (IDs) y métricas numéricas agregables (`Monto Total Vigente`, `Monto del Donante`, `MontoDolarizado`, `Monto Subdonación`). No deben contener textos descriptivos.

4.  **Tablas Bridge (M:N)**:
    Las relaciones de muchos a muchos (`Bridge_Proyecto_Ambito`, `Bridge_Proyecto_Ubicacion`, `Bridge_Subdonacion_Ubicacion`) se manejan expandiendo las celdas concatenadas por `;` mediante un `explode()`.
    *   *Nota para Power BI*: La dimensión principal filtra la Bridge de manera bidireccional (Ambos), pero la dimensión satélite filtra la Bridge de forma única.

## Reglas de Procesamiento de Datos (Python ETL)

Al editar o interactuar con el script `script/v7_ET.py`, debes adherirte a las siguientes reglas de negocio:

*   **Estandarización de Códigos (Padding)**:
    Los IDs o códigos (como `Código` o `Identificador entidad`) suelen perder sus ceros a la izquierda al exportarse desde la base de datos origen. El código en Python debe limpiar estos campos y aplicarles un *Zero-Padding*: 3 dígitos si el valor original tiene 3 o menos (ej. `45` -> `045`), y 6 dígitos si tiene más de 3 (ej. `45090` -> `045090`).
*   **Dolarización Basada en Diccionarios**:
    Todos los montos aportados por donantes deben ser dolarizados para homologación. Nunca utilices lógicas *hardcoded* dispersas (ej. múltiples condicionales IF). Utiliza el diccionario centralizado `FACTOR_A_USD` definido en el script. Si surge una nueva moneda (ej. Franco Suizo, Dólar Canadiense), simplemente agrega su factor de conversión a este diccionario.

## Estructura de Directorios

*   `data/inputs/`: Contiene los tres archivos Excel crudos (`Profonanpe - entidades.xlsx`, `Profonanpe - subproyectos.xlsx` y `Profonanpe - registros.xlsx`).
*   `data/outputs/`: Contiene el archivo Excel finalizado (`modelo_estrella_proyectos.xlsx`) que alimenta directamente el dashboard de Power BI.
*   `script/`: Contiene el script principal de transformación (`v7_ET.py`).

**Nota para el Agente:** Siempre valida que no haya archivos de Excel abiertos por el usuario antes de correr los scripts, ya que esto disparará errores de bloqueo de sistema de archivos en Windows (Permission Denied).
