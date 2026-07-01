import pandas as pd
import numpy as np

# =========================================================
# 1. LEER ARCHIVOS
# =========================================================

df_proyectos = pd.read_excel(
    "Profonanpe -  entidades.xlsx"
)

df_sub = pd.read_excel(
    "Profonanpe -  subproyectos.xlsx"
)

df_informes = pd.read_excel(
    "Profonanpe - registros.xlsx"
)

# =========================================================
# 2. LIMPIAR COLUMNAS
# =========================================================

df_proyectos.columns = (
    df_proyectos.columns.str.strip()
)

df_sub.columns = (
    df_sub.columns.str.strip()
)

df_informes.columns = (
    df_informes.columns.str.strip()
)

# =========================================================
# 3. FACT PROYECTOS
# =========================================================

fact_proyectos = df_proyectos.drop(
    columns=[
        'Departamento',
        'Rol de Profonanpe',
        'Eje Temático',
        'Línea de Negocio',

        # Columnas sin análisis.
        'Subdonaciones',
        'Eje temático',
        'Año de inicio',
        'Adenda',
        'Tipo',
        'Ubicación / Distritos (identificador)',
        'Ubicación / Distritos (alias)',
        'Localidad / Comunidades (identificador)',
        'Localidad / Comunidades (alias)',
        'Tipo de cambio',
        'Convenio',
        'Adenda (si existe)',
        'Punto Focal  DIME',
        'Articulación  con Entidades Académicas ( Descripción)',
        'Documentos Adicionales (Drive)',
        'Monto de la contrapartida'
    ],
    errors='ignore'
).copy()

fact_proyectos = (
    fact_proyectos
    .drop_duplicates()
    .reset_index(drop=True)
)

# Renombrar
fact_proyectos = fact_proyectos.rename(columns={
    'Código': 'CodigoProyecto'
})

# ID surrogate
fact_proyectos['IdProyecto'] = (
    fact_proyectos.index + 1
)

# =========================================================
# 3.1 FACT INFORMES
# =========================================================

fact_informes = df_informes.copy()

fact_informes = fact_informes[[
    'Fecha',
    'Estado',
    'Autor',
    'Identificador entidad',
    'Tipo de Informe',
    'Periodicidad',
    'Fecha de presentación según convenio',
    'Fecha de presentación real',
    'Estado.1',
    'Aprobado'
]].copy()

fact_informes = fact_informes.rename(columns={
    'Fecha': 'FechaRegistro',
    'Estado': 'EstadoWorkflow',
    'Identificador entidad': 'CodigoProyecto',
    'Fecha de presentación según convenio': 'FechaPresentacionConvenio',
    'Fecha de presentación real': 'FechaPresentacionReal',
    'Estado.1': 'EstadoInforme'
})

# Relacionar con proyecto
fact_informes = (
    fact_informes.merge(
        fact_proyectos[
            ['CodigoProyecto', 'IdProyecto']
        ],
        on='CodigoProyecto',
        how='left'
    )
)

# Fecha Registro a dd-mm-yyyy
fact_informes['FechaRegistro'] = (
    pd.to_datetime(
        fact_informes['FechaRegistro'],
        errors='coerce'
    )
    .dt.strftime('%d-%m-%Y')
)

# Id surrogate
fact_informes['IdInforme'] = (
    fact_informes.index + 1
)

# =========================================================
# 4. TIPO DE CAMBIO Y DOLARIZACIÓN
# =========================================================

tipo_cambio_usd = 3.394
tipo_cambio_eur = 3.962176

fact_proyectos['TipoCambioUSD'] = (
    tipo_cambio_usd
)

fact_proyectos['TipoCambioEUR'] = (
    tipo_cambio_eur
)

# Convertir numérico
fact_proyectos['Monto del donante'] = pd.to_numeric(
    fact_proyectos['Monto del donante'],
    errors='coerce'
)

# Inicializar
fact_proyectos['MontoDolarizado'] = np.nan

# SOL → USD
mask_soles = (
    fact_proyectos['Moneda']
    .astype(str)
    .str.upper()
    .isin(['SOLES', 'PEN'])
)

fact_proyectos.loc[
    mask_soles,
    'MontoDolarizado'
] = (
    fact_proyectos.loc[
        mask_soles,
        'Monto del donante'
    ] / tipo_cambio_usd
)

# EUR → USD
mask_euros = (
    fact_proyectos['Moneda']
    .astype(str)
    .str.upper()
    .isin(['EUROS', 'EUR'])
)

fact_proyectos.loc[
    mask_euros,
    'MontoDolarizado'
] = (
    fact_proyectos.loc[
        mask_euros,
        'Monto del donante'
    ] * tipo_cambio_eur / tipo_cambio_usd
)

# USD → USD
mask_usd = (
    fact_proyectos['Moneda']
    .astype(str)
    .str.upper()
    .isin(['DÓLARES', 'DOLARES', 'USD'])
)

fact_proyectos.loc[
    mask_usd,
    'MontoDolarizado'
] = (
    fact_proyectos.loc[
        mask_usd,
        'Monto del donante'
    ]
)

# Redondeo
fact_proyectos['MontoDolarizado'] = (
    fact_proyectos['MontoDolarizado']
    .astype(float)
    .round(2)
)

# =========================================================
# 5. DIMENSION UBICACION
# =========================================================

map_departamentos = {
    'AMAZONAS': '01',
    'ANCASH': '02',
    'APURIMAC': '03',
    'AREQUIPA': '04',
    'AYACUCHO': '05',
    'CAJAMARCA': '06',
    'CALLAO': '07',
    'CUSCO': '08',
    'HUANCAVELICA': '09',
    'HUANUCO': '10',
    'ICA': '11',
    'JUNIN': '12',
    'LA LIBERTAD': '13',
    'LAMBAYEQUE': '14',
    'LIMA': '15',
    'LORETO': '16',
    'MADRE DE DIOS': '17',
    'MOQUEGUA': '18',
    'PASCO': '19',
    'PIURA': '20',
    'PUNO': '21',
    'SAN MARTIN': '22',
    'TACNA': '23',
    'TUMBES': '24',
    'UCAYALI': '25'
}

# =========================
# UBICACIONES PROYECTOS
# =========================

ubic_proy = df_proyectos[
    ['Departamento']
].dropna().copy()

ubic_proy['Departamento'] = (
    ubic_proy['Departamento']
    .str.split(';')
)

ubic_proy = (
    ubic_proy.explode('Departamento')
)

ubic_proy['Departamento'] = (
    ubic_proy['Departamento']
    .str.strip()
    .str.upper()
)

ubic_proy['Departamento'] = (
    ubic_proy['Departamento']
    .str.normalize('NFKD')
    .str.encode('ascii', errors='ignore')
    .str.decode('utf-8')
)

ubic_proy['Ubigeo'] = (
    ubic_proy['Departamento']
    .map(map_departamentos)
)

ubic_proy['Provincia'] = np.nan
ubic_proy['Distrito'] = np.nan

ubic_proy['NivelUbicacion'] = (
    'Departamento'
)

# =========================
# UBICACIONES SUBDONACIONES
# =========================

ubic_sub = df_sub[[
    'Ubicación (identificador)',
    'Ubicación (alias)'
]].dropna().copy()

ubic_sub = ubic_sub.rename(columns={
    'Ubicación (identificador)': 'Ubigeo',
    'Ubicación (alias)': 'UbicacionAlias'
})

ubic_sub['Ubigeo'] = (
    ubic_sub['Ubigeo']
    .fillna('')
    .astype(str)
    .str.replace('.0', '', regex=False)
    .str.zfill(6)
)

split_geo = (
    ubic_sub['UbicacionAlias']
    .str.split(',', expand=True)
)

ubic_sub['Departamento'] = (
    split_geo[0]
    .str.strip()
    .str.upper()
)

ubic_sub['Provincia'] = (
    split_geo[1]
    .str.strip()
    .str.upper()
)

ubic_sub['Distrito'] = (
    split_geo[2]
    .str.strip()
    .str.upper()
)

ubic_sub['NivelUbicacion'] = (
    'Distrito'
)

# =========================
# UNIFICAR
# =========================

dim_ubicacion = pd.concat([
    ubic_proy[[
        'Ubigeo',
        'Departamento',
        'Provincia',
        'Distrito',
        'NivelUbicacion'
    ]],
    ubic_sub[[
        'Ubigeo',
        'Departamento',
        'Provincia',
        'Distrito',
        'NivelUbicacion'
    ]]
], ignore_index=True)

dim_ubicacion = (
    dim_ubicacion
    .drop_duplicates()
    .reset_index(drop=True)
)

dim_ubicacion['IdUbicacion'] = (
    dim_ubicacion.index + 1
)

dim_ubicacion['Ubigeo'] = (
    dim_ubicacion['Ubigeo']
    .astype(str)
    .str.strip()
)

# =========================================================
# 6. BRIDGE PROYECTO UBICACION
# =========================================================

bridge_proyecto_ubicacion = df_proyectos[
    ['Código', 'Departamento']
].dropna().copy()

bridge_proyecto_ubicacion['Departamento'] = (
    bridge_proyecto_ubicacion['Departamento']
    .str.split(';')
)

bridge_proyecto_ubicacion = (
    bridge_proyecto_ubicacion
    .explode('Departamento')
)

bridge_proyecto_ubicacion['Departamento'] = (
    bridge_proyecto_ubicacion['Departamento']
    .str.strip()
    .str.upper()
)

bridge_proyecto_ubicacion['Departamento'] = (
    bridge_proyecto_ubicacion['Departamento']
    .str.normalize('NFKD')
    .str.encode('ascii', errors='ignore')
    .str.decode('utf-8')
)

bridge_proyecto_ubicacion['Ubigeo'] = (
    bridge_proyecto_ubicacion['Departamento']
    .map(map_departamentos)
)

bridge_proyecto_ubicacion = (
    bridge_proyecto_ubicacion.rename(columns={
        'Código': 'CodigoProyecto'
    })
)

bridge_proyecto_ubicacion = (
    bridge_proyecto_ubicacion.merge(
        fact_proyectos[
            ['CodigoProyecto', 'IdProyecto']
        ],
        on='CodigoProyecto',
        how='left'
    )
)

bridge_proyecto_ubicacion = (
    bridge_proyecto_ubicacion.merge(
        dim_ubicacion[
            ['Ubigeo', 'IdUbicacion']
        ],
        on='Ubigeo',
        how='left'
    )
)

bridge_proyecto_ubicacion = (
    bridge_proyecto_ubicacion[
        ['IdProyecto', 'IdUbicacion']
    ]
    .drop_duplicates()
)

# =========================================================
# 7. DIMENSION ROLES
# =========================================================

puente_rol = df_proyectos[[
    'Código',
    'Rol de Profonanpe'
]].dropna().copy()

puente_rol['Rol de Profonanpe'] = (
    puente_rol['Rol de Profonanpe']
    .str.strip()
)

puente_rol['Rol de Profonanpe'] = (
    puente_rol['Rol de Profonanpe']
    .str.split(';')
)

puente_rol['Rol de Profonanpe'] = (
    puente_rol['Rol de Profonanpe']
    .apply(
        lambda x: sorted(
            [i.strip().title() for i in x]
        )
    )
)

puente_rol['Rol de Profonanpe'] = (
    puente_rol['Rol de Profonanpe']
    .str.join(' y ')
)

puente_rol = puente_rol.drop_duplicates()

dim_roles = (
    puente_rol[['Rol de Profonanpe']]
    .drop_duplicates()
    .reset_index(drop=True)
)

dim_roles['IdRol'] = (
    dim_roles.index + 1
)

puente_rol = puente_rol.rename(columns={
    'Código': 'CodigoProyecto'
})

puente_rol = puente_rol.merge(
    dim_roles,
    on='Rol de Profonanpe',
    how='left'
)

puente_rol = puente_rol.merge(
    fact_proyectos[
        ['CodigoProyecto', 'IdProyecto']
    ],
    on='CodigoProyecto',
    how='left'
)

bridge_roles = puente_rol[[
    'IdProyecto',
    'IdRol'
]].drop_duplicates()

# =========================================================
# 8. DIMENSION SUBDONACION
# =========================================================

dim_subdonacion = df_sub.copy()

# =========================================================
# RENOMBRAR
# =========================================================

dim_subdonacion = dim_subdonacion.rename(columns={
    'Código': 'CodigoSubdonacion',
    'Tipo de subdonación': 'TipoSubdonacion',
    'Proyecto Nivel 0 relacionado (identificador)': 'CodigoProyecto',
    'Proyecto nivel 1 (identificador)': 'CodigoSubNivel1'
})

# =========================================================
# ID SURROGATE
# =========================================================

dim_subdonacion['IdSubdonacion'] = (
    dim_subdonacion.index + 1
)

# =========================================================
# RELACIONAR PROYECTO
# =========================================================

dim_subdonacion = (
    dim_subdonacion.merge(
        fact_proyectos[
            ['CodigoProyecto', 'IdProyecto']
        ],
        on='CodigoProyecto',
        how='left'
    )
)

# =========================================================
# NIVEL SUBDONACION
# =========================================================

dim_subdonacion['NivelSubdonacion'] = np.nan

mask_n1 = (
    dim_subdonacion['TipoSubdonacion']
    .astype(str)
    .str.upper()
    .str.contains('NIVEL 1')
)

mask_n2 = (
    dim_subdonacion['TipoSubdonacion']
    .astype(str)
    .str.upper()
    .str.contains('NIVEL 2')
)

dim_subdonacion.loc[
    mask_n1,
    'NivelSubdonacion'
] = 1

dim_subdonacion.loc[
    mask_n2,
    'NivelSubdonacion'
] = 2

# =========================================================
# MAPEAR CODIGO NIVEL 1 → ID NIVEL 1
# =========================================================

map_nivel1 = (
    dim_subdonacion[
        dim_subdonacion['NivelSubdonacion'] == 1
    ][[
        'CodigoSubdonacion',
        'IdSubdonacion'
    ]]
    .drop_duplicates()
)

map_nivel1.columns = [
    'CodigoSubNivel1',
    'IdSubdonacionPadre'
]

# =========================================================
# RELACIONAR PADRE
# =========================================================

dim_subdonacion['CodigoSubNivel1'] = (
    dim_subdonacion['CodigoSubNivel1']
    .astype(str)
)

map_nivel1['CodigoSubNivel1'] = (
    map_nivel1['CodigoSubNivel1']
    .astype(str)
)

dim_subdonacion = (
    dim_subdonacion.merge(
        map_nivel1,
        on='CodigoSubNivel1',
        how='left'
    )
)

# Nivel 1 no tiene padre
dim_subdonacion.loc[
    dim_subdonacion['NivelSubdonacion'] == 1,
    'IdSubdonacionPadre'
] = np.nan

# =========================================================
# 9. BRIDGE SUBDONACION UBICACION
# =========================================================

bridge_subdonacion_ubicacion = (
    dim_subdonacion.copy()
)

bridge_subdonacion_ubicacion = (
    bridge_subdonacion_ubicacion.rename(columns={
        'Ubicación (identificador)': 'Ubigeo'
    })
)

bridge_subdonacion_ubicacion['Ubigeo'] = (
    bridge_subdonacion_ubicacion['Ubigeo']
    .fillna('')
    .astype(str)
    .str.replace('.0', '', regex=False)
    .str.zfill(6)
)

bridge_subdonacion_ubicacion = (
    bridge_subdonacion_ubicacion.merge(
        dim_ubicacion[
            ['Ubigeo', 'IdUbicacion']
        ],
        on='Ubigeo',
        how='left'
    )
)

bridge_subdonacion_ubicacion = (
    bridge_subdonacion_ubicacion[[
        'IdSubdonacion',
        'IdUbicacion'
    ]]
    .drop_duplicates()
)

# =========================================================
# 10. DIMENSION EJE TEMATICO
# =========================================================

puente_eje = df_sub[
    ['Código', 'Eje Temático']
].dropna().copy()

puente_eje['Eje Temático'] = (
    puente_eje['Eje Temático']
    .astype(str)
    .str.split(';')
)

puente_eje = (
    puente_eje.explode('Eje Temático')
)

puente_eje['Eje Temático'] = (
    puente_eje['Eje Temático']
    .str.strip()
)

# DIMENSION

dim_eje_tematico = (
    puente_eje[['Eje Temático']]
    .drop_duplicates()
    .sort_values('Eje Temático')
    .reset_index(drop=True)
)

dim_eje_tematico['IdEjeTematico'] = (
    dim_eje_tematico.index + 1
)

# BRIDGE

bridge_subdonacion_eje = (
    puente_eje.rename(columns={
        'Código': 'CodigoSubdonacion'
    })
)

bridge_subdonacion_eje = (
    bridge_subdonacion_eje.merge(
        dim_eje_tematico,
        on='Eje Temático',
        how='left'
    )
)

bridge_subdonacion_eje = (
    bridge_subdonacion_eje.merge(
        dim_subdonacion[
            ['CodigoSubdonacion', 'IdSubdonacion']
        ],
        on='CodigoSubdonacion',
        how='left'
    )
)

bridge_subdonacion_eje = (
    bridge_subdonacion_eje[
        ['IdSubdonacion', 'IdEjeTematico']
    ]
    .drop_duplicates()
)

# =========================================================
# 11. DIMENSION LINEA NEGOCIO
# =========================================================

puente_linea = df_sub[
    ['Código', 'Línea de Negocio']
].dropna().copy()

puente_linea['Línea de Negocio'] = (
    puente_linea['Línea de Negocio']
    .astype(str)
    .str.split(';')
)

puente_linea = (
    puente_linea.explode('Línea de Negocio')
)

puente_linea['Línea de Negocio'] = (
    puente_linea['Línea de Negocio']
    .str.strip()
)

# DIMENSION

dim_linea_negocio = (
    puente_linea[['Línea de Negocio']]
    .drop_duplicates()
    .sort_values('Línea de Negocio')
    .reset_index(drop=True)
)

dim_linea_negocio['IdLineaNegocio'] = (
    dim_linea_negocio.index + 1
)

# BRIDGE

bridge_subdonacion_linea = (
    puente_linea.rename(columns={
        'Código': 'CodigoSubdonacion'
    })
)

bridge_subdonacion_linea = (
    bridge_subdonacion_linea.merge(
        dim_linea_negocio,
        on='Línea de Negocio',
        how='left'
    )
)

bridge_subdonacion_linea = (
    bridge_subdonacion_linea.merge(
        dim_subdonacion[
            ['CodigoSubdonacion', 'IdSubdonacion']
        ],
        on='CodigoSubdonacion',
        how='left'
    )
)

bridge_subdonacion_linea = (
    bridge_subdonacion_linea[
        ['IdSubdonacion', 'IdLineaNegocio']
    ]
    .drop_duplicates()
)

# =========================================================
# LIMPIAR DIM_SUBDONACION
# =========================================================

dim_subdonacion = dim_subdonacion.drop(
    columns=[
        'Año inicio',
        'Adenda',
        'Monto de Contrapartida (PEN)',
        'Monto Sinergias y Otros',
        'Ubicación web',
        'Convenio subdonación',
        'Adenda (si existe)',
        'Monto del donante',
        'Documentos Adicionales al Convenio',
        'Enlace de archivo del poligono de la IL',
        'Enlace de Acta de Comité Local establecido',

        # Ya representadas en Dim_Ubicacion + Bridge
        'Ubicación (identificador)',
        'Ubicación (alias)',
        'Localidad (identificador)',
        'Localidad (alias)',

        # Ya representadas en Dim_EjeTematico y Dim_LineaNegocio
        'Eje Temático',
        'Línea de Negocio'
    ],
    errors='ignore'
)

# =========================================================
# 12. EXPORTAR
# =========================================================

with pd.ExcelWriter(
    "modelo_estrella_proyectos.xlsx",
    engine="openpyxl"
) as writer:

    fact_proyectos.to_excel(
        writer,
        sheet_name="Fact_Proyectos",
        index=False
    )

    fact_informes.to_excel(
        writer,
        sheet_name="Fact_Informes",
        index=False
    )
    
    dim_ubicacion.to_excel(
        writer,
        sheet_name="Dim_Ubicacion",
        index=False
    )

    bridge_proyecto_ubicacion.to_excel(
        writer,
        sheet_name="Bridge_Proyecto_Ubicacion",
        index=False
    )

    dim_roles.to_excel(
        writer,
        sheet_name="Dim_Roles",
        index=False
    )

    bridge_roles.to_excel(
        writer,
        sheet_name="Bridge_Proyecto_Rol",
        index=False
    )

    dim_subdonacion.to_excel(
        writer,
        sheet_name="Dim_Subdonacion",
        index=False
    )

    bridge_subdonacion_ubicacion.to_excel(
        writer,
        sheet_name="Bridge_Subdonacion_Ubicacion",
        index=False
    )

    dim_eje_tematico.to_excel(
        writer,
        sheet_name="Dim_EjeTematico",
        index=False
    )

    bridge_subdonacion_eje.to_excel(
        writer,
        sheet_name="Bridge_Subdonacion_EjeTematico",
        index=False
    )

    dim_linea_negocio.to_excel(
        writer,
        sheet_name="Dim_LineaNegocio",
        index=False
    )

    bridge_subdonacion_linea.to_excel(
        writer,
        sheet_name="Bridge_Subdonacion_LineaNegocio",
        index=False
    )

print("Modelo estrella v5 generado correctamente")