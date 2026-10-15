import pandas as pd
import numpy as np

def format_codigo(val):
    if pd.isna(val): return val
    s = str(val).replace('.0', '').strip()
    if len(s) <= 3: return s.zfill(3)
    return s.zfill(6)


# =========================================================
# 1. LEER ARCHIVOS
# =========================================================

df_proyectos = pd.read_excel(r'../data/Input/Profonanpe - entidades.xlsx')
df_sub = pd.read_excel(r'../data/Input/Profonanpe - subproyectos.xlsx')
df_informes = pd.read_excel(r'../data/Input/Profonanpe - registros.xlsx')

# =========================================================
# 2. LIMPIAR COLUMNAS
# =========================================================

df_proyectos.columns = df_proyectos.columns.str.strip()
df_sub.columns = df_sub.columns.str.strip()
df_informes.columns = df_informes.columns.str.strip()

df_proyectos['Código'] = df_proyectos['Código'].dropna().astype(str).str.replace('.0', '', regex=False).str.strip()
df_proyectos['Código'] = df_proyectos['Código'].apply(format_codigo)

# =========================================================
# 4. DIM_PROYECTOS (Textos y Descriptivos)
# =========================================================

dim_proyectos = df_proyectos[[
    'Código', 'Nombre', 'Nombre Corto', 'Fecha de Inicio', 'Fecha de Cierre', 'Nueva Fecha de Cierre',
    'Seguimiento', 'Objetivo General', 'Beneficiarios', 
    'Aliados para la Ejecución (Gores, Municipalidades y  Empresas)', 
    'Avances Clave ( al 31.12.25)', 'Ejecutores', 'Donantes', 'Moneda', 
    'Jefe o Coordinador', 'Estado Plan Operativo Anual', 'Rol de PROFONANPE', 'Área Temática', 'Tipo de Intervención', 'Tipo Recurso Global'
]].copy()

dim_proyectos = dim_proyectos.drop_duplicates().reset_index(drop=True)
dim_proyectos = dim_proyectos.rename(columns={'Código': 'CodigoProyecto'})
dim_proyectos['IdProyecto'] = dim_proyectos.index + 1

# =========================================================
# 4.1 FACT PROYECTOS (Solo Métricas y FKs)
# =========================================================

fact_proyectos = df_proyectos[[
    'Código', 
    'Monto Total Vigente', 'Monto del Donante', 'Moneda'
]].copy()

fact_proyectos = fact_proyectos.rename(columns={'Código': 'CodigoProyecto'})
fact_proyectos = fact_proyectos.merge(dim_proyectos[['CodigoProyecto', 'IdProyecto']], on='CodigoProyecto', how='left')

# =========================================================
# 4.2 FACT INFORMES
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

fact_informes['Identificador entidad'] = fact_informes['Identificador entidad'].apply(format_codigo)
fact_informes = fact_informes.rename(columns={
    'Fecha': 'FechaRegistro',
    'Estado': 'EstadoWorkflow',
    'Identificador entidad': 'CodigoProyecto',
    'Fecha de presentación según convenio': 'FechaPresentacionConvenio',
    'Fecha de presentación real': 'FechaPresentacionReal',
    'Estado.1': 'EstadoInforme'
})

fact_informes = fact_informes.merge(
    dim_proyectos[['CodigoProyecto', 'IdProyecto']],
    on='CodigoProyecto',
    how='left'
)

fact_informes['FechaRegistro'] = pd.to_datetime(fact_informes['FechaRegistro'], errors='coerce').dt.strftime('%d-%m-%Y')
fact_informes['IdInforme'] = fact_informes.index + 1

# =========================================================
# 5. TIPO DE CAMBIO Y DOLARIZACIÓN
# =========================================================
# Diccionario maestro: factor de conversión directa a USD.
#   valor = cuántos USD vale 1 unidad de esa moneda.
#   Para agregar una moneda nueva, solo añadir una línea aquí.

FACTOR_A_USD = {
    # Formato nuevo (Post - Profonanpe - entidades)
    'Dólar estadounidense (USD)':  1.0,
    'Sol peruano (PEN)':           1/3.4, 
    'Euro (EUR)':                  1.1,
    'Dólar canadiense (CAD)':      0.72,
    'Dólar australiano (AUD)':     0.66,
    'Libra esterlina (GBP)':       1.27,
    'Franco suizo (CHF)':          1.14,
    # Formato legacy (Previo - Profonanpe - entidades)
    'SOLES':   1 / 3.394,
    'PEN':     1 / 3.394,
    'EUROS':   1.08,
    'EUR':     1.08,
    'DÓLARES': 1.0,
    'DOLARES': 1.0,
    'USD':     1.0,
}

# Mapear cada proyecto a su factor de conversión
fact_proyectos['Monto del Donante'] = pd.to_numeric(fact_proyectos['Monto del Donante'], errors='coerce')
fact_proyectos['Monto Total Vigente'] = pd.to_numeric(fact_proyectos['Monto Total Vigente'], errors='coerce')

fact_proyectos['FactorUSD'] = (
    fact_proyectos['Moneda']
    .astype(str)
    .str.strip()
    .map(FACTOR_A_USD)
)

fact_proyectos['MontoDolarizado'] = (
    fact_proyectos['Monto del Donante'] * fact_proyectos['FactorUSD']
).round(2)

# Alertar si hay monedas no mapeadas
sin_mapear = fact_proyectos[fact_proyectos['FactorUSD'].isna() & fact_proyectos['Moneda'].notna()]['Moneda'].unique()
if len(sin_mapear) > 0:
    print(f"  Monedas sin tipo de cambio (MontoDolarizado = NaN): {list(sin_mapear)}")

# Limpiar Fact_Proyectos: Moneda vive en Dim_Proyectos
fact_proyectos = fact_proyectos.drop(columns=['Moneda', 'CodigoProyecto', 'FactorUSD'])
fact_proyectos = fact_proyectos[['IdProyecto', 'Monto Total Vigente', 'Monto del Donante', 'MontoDolarizado']]

# =========================================================
# 6. DIMENSION AMBITO
# =========================================================

bridge_ambito = df_proyectos[['Código', 'Ámbito de Intervención']].dropna().copy()
bridge_ambito = bridge_ambito.rename(columns={'Código': 'CodigoProyecto'})

bridge_ambito['Ámbito de Intervención'] = bridge_ambito['Ámbito de Intervención'].astype(str).str.split(';')
bridge_ambito = bridge_ambito.explode('Ámbito de Intervención')
bridge_ambito['Ámbito de Intervención'] = bridge_ambito['Ámbito de Intervención'].str.strip()

dim_ambito = bridge_ambito[['Ámbito de Intervención']].drop_duplicates().reset_index(drop=True)
dim_ambito['IdAmbito'] = dim_ambito.index + 1

bridge_ambito = bridge_ambito.merge(dim_ambito, on='Ámbito de Intervención', how='left')
bridge_ambito = bridge_ambito.merge(dim_proyectos[['CodigoProyecto', 'IdProyecto']], on='CodigoProyecto', how='left')
bridge_proyecto_ambito = bridge_ambito[['IdProyecto', 'IdAmbito']].drop_duplicates()

# =========================================================
# 7. DIMENSION UBICACION
# =========================================================

map_departamentos = {
    'AMAZONAS': '01', 'ANCASH': '02', 'APURIMAC': '03', 'AREQUIPA': '04', 'AYACUCHO': '05',
    'CAJAMARCA': '06', 'CALLAO': '07', 'CUSCO': '08', 'HUANCAVELICA': '09', 'HUANUCO': '10',
    'ICA': '11', 'JUNIN': '12', 'LA LIBERTAD': '13', 'LAMBAYEQUE': '14', 'LIMA': '15',
    'LORETO': '16', 'MADRE DE DIOS': '17', 'MOQUEGUA': '18', 'PASCO': '19', 'PIURA': '20',
    'PUNO': '21', 'SAN MARTIN': '22', 'TACNA': '23', 'TUMBES': '24', 'UCAYALI': '25'
}

ubic_proy = df_proyectos[['Departamentos']].dropna().copy()
ubic_proy['Departamentos'] = ubic_proy['Departamentos'].str.split(';')
ubic_proy = ubic_proy.explode('Departamentos')
ubic_proy['Departamentos'] = ubic_proy['Departamentos'].str.strip().str.upper()
ubic_proy['Departamentos'] = ubic_proy['Departamentos'].str.normalize('NFKD').str.encode('ascii', errors='ignore').str.decode('utf-8')
ubic_proy['Ubigeo'] = ubic_proy['Departamentos'].map(map_departamentos)
ubic_proy['Departamento'] = ubic_proy['Departamentos']
ubic_proy['Provincia'] = np.nan
ubic_proy['Distrito'] = np.nan
ubic_proy['NivelUbicacion'] = 'Departamento'

ubic_sub = df_sub[['Ubicación / Distritos (identificador)', 'Ubicación / Distritos (alias)']].dropna().copy()
ubic_sub = ubic_sub.rename(columns={'Ubicación / Distritos (identificador)': 'Ubigeo', 'Ubicación / Distritos (alias)': 'UbicacionAlias'})
ubic_sub['Ubigeo'] = ubic_sub['Ubigeo'].fillna('').astype(str).str.replace('.0', '', regex=False)

ubic_sub['Ubigeo'] = ubic_sub['Ubigeo'].str.split(';')
ubic_sub['UbicacionAlias'] = ubic_sub['UbicacionAlias'].astype(str).str.split(';')
ubic_sub = ubic_sub.explode(['Ubigeo', 'UbicacionAlias'])
ubic_sub['Ubigeo'] = ubic_sub['Ubigeo'].astype(str).str.strip().str.zfill(6)
ubic_sub['UbicacionAlias'] = ubic_sub['UbicacionAlias'].astype(str).str.strip()

split_geo = ubic_sub['UbicacionAlias'].str.split(',', expand=True)
ubic_sub['Departamento'] = split_geo[0].str.strip().str.upper() if split_geo.shape[1] > 0 else np.nan
ubic_sub['Provincia'] = split_geo[1].str.strip().str.upper() if split_geo.shape[1] > 1 else np.nan
ubic_sub['Distrito'] = split_geo[2].str.strip().str.upper() if split_geo.shape[1] > 2 else np.nan
ubic_sub['NivelUbicacion'] = 'Distrito'

# DIMENSION UBICACION PROYECTO (Solo Departamentos)
dim_ubicacion_proyecto = ubic_proy[['Ubigeo', 'Departamento']].drop_duplicates().reset_index(drop=True)
dim_ubicacion_proyecto['IdUbicacion'] = dim_ubicacion_proyecto.index + 1
dim_ubicacion_proyecto['Ubigeo'] = dim_ubicacion_proyecto['Ubigeo'].astype(str).str.strip()

# DIMENSION UBICACION SUBDONACION (Distritos, Provincias y Departamentos)
dim_ubicacion_subdonacion = ubic_sub[['Ubigeo', 'Departamento', 'Provincia', 'Distrito']].drop_duplicates().reset_index(drop=True)
dim_ubicacion_subdonacion['IdUbicacion'] = dim_ubicacion_subdonacion.index + 1
dim_ubicacion_subdonacion['Ubigeo'] = dim_ubicacion_subdonacion['Ubigeo'].astype(str).str.strip()

bridge_proyecto_ubicacion = df_proyectos[['Código', 'Departamentos']].dropna().copy()
bridge_proyecto_ubicacion['Departamentos'] = bridge_proyecto_ubicacion['Departamentos'].str.split(';')
bridge_proyecto_ubicacion = bridge_proyecto_ubicacion.explode('Departamentos')
bridge_proyecto_ubicacion['Departamentos'] = bridge_proyecto_ubicacion['Departamentos'].str.strip().str.upper()
bridge_proyecto_ubicacion['Departamentos'] = bridge_proyecto_ubicacion['Departamentos'].str.normalize('NFKD').str.encode('ascii', errors='ignore').str.decode('utf-8')
bridge_proyecto_ubicacion['Ubigeo'] = bridge_proyecto_ubicacion['Departamentos'].map(map_departamentos)
bridge_proyecto_ubicacion = bridge_proyecto_ubicacion.rename(columns={'Código': 'CodigoProyecto'})
bridge_proyecto_ubicacion = bridge_proyecto_ubicacion.merge(dim_proyectos[['CodigoProyecto', 'IdProyecto']], on='CodigoProyecto', how='left')
bridge_proyecto_ubicacion = bridge_proyecto_ubicacion.merge(dim_ubicacion_proyecto[['Ubigeo', 'IdUbicacion']], on='Ubigeo', how='left')
bridge_proyecto_ubicacion = bridge_proyecto_ubicacion.dropna(subset=['IdUbicacion'])[['IdProyecto', 'IdUbicacion']].drop_duplicates()

# =========================================================
# 8. DIMENSION SUBDONACION
# =========================================================

dim_subdonacion = df_sub.copy()
dim_subdonacion['Intervención Origen (identificador)'] = dim_subdonacion['Intervención Origen (identificador)'].apply(format_codigo)
dim_subdonacion = dim_subdonacion.rename(columns={
    'Código': 'CodigoSubdonacion',
    'Intervención Origen (identificador)': 'CodigoProyecto',
    'Intervención Origen (alias)': 'ProyectoOrigenAlias',
    'Subdonación Origen (identificador)': 'CodigoSubNivel1'
})
dim_subdonacion['IdSubdonacion'] = dim_subdonacion.index + 1
dim_subdonacion = dim_subdonacion.merge(dim_proyectos[['CodigoProyecto', 'IdProyecto']], on='CodigoProyecto', how='left')

# Jerarquía: Nivel 1 vs Nivel 2 dinámico
dim_subdonacion['NivelSubdonacion'] = 1
mask_n2 = dim_subdonacion['CodigoSubNivel1'].notna() & (dim_subdonacion['CodigoSubNivel1'].astype(str).str.strip() != '') & (dim_subdonacion['CodigoSubNivel1'].astype(str).str.lower() != 'nan')
dim_subdonacion.loc[mask_n2, 'NivelSubdonacion'] = 2

map_nivel1 = dim_subdonacion[dim_subdonacion['NivelSubdonacion'] == 1][['CodigoSubdonacion', 'IdSubdonacion']].drop_duplicates()
map_nivel1.columns = ['CodigoSubNivel1', 'IdSubdonacionPadre']

dim_subdonacion['CodigoSubNivel1'] = dim_subdonacion['CodigoSubNivel1'].astype(str)
map_nivel1['CodigoSubNivel1'] = map_nivel1['CodigoSubNivel1'].astype(str)
dim_subdonacion = dim_subdonacion.merge(map_nivel1, on='CodigoSubNivel1', how='left')
dim_subdonacion.loc[dim_subdonacion['NivelSubdonacion'] == 1, 'IdSubdonacionPadre'] = np.nan

# Bridge Subdonacion Ubicacion
bridge_subdonacion_ubicacion = dim_subdonacion.copy()
bridge_subdonacion_ubicacion = bridge_subdonacion_ubicacion.rename(columns={'Ubicación / Distritos (identificador)': 'Ubigeo'})
bridge_subdonacion_ubicacion['Ubigeo'] = bridge_subdonacion_ubicacion['Ubigeo'].fillna('').astype(str).str.replace('.0', '', regex=False)

bridge_subdonacion_ubicacion['Ubigeo'] = bridge_subdonacion_ubicacion['Ubigeo'].str.split(';')
bridge_subdonacion_ubicacion = bridge_subdonacion_ubicacion.explode('Ubigeo')
bridge_subdonacion_ubicacion['Ubigeo'] = bridge_subdonacion_ubicacion['Ubigeo'].astype(str).str.strip().str.zfill(6)

bridge_subdonacion_ubicacion = bridge_subdonacion_ubicacion.merge(dim_ubicacion_subdonacion[['Ubigeo', 'IdUbicacion']], on='Ubigeo', how='left')
bridge_subdonacion_ubicacion = bridge_subdonacion_ubicacion.dropna(subset=['IdUbicacion'])[['IdSubdonacion', 'IdUbicacion']].drop_duplicates()

# =========================================================
# 9. FACT SUBDONACION (Métricas y FKs)
# =========================================================

fact_subdonacion = dim_subdonacion[[
    'IdSubdonacion', 'IdProyecto', 'IdSubdonacionPadre', 'NivelSubdonacion',
    'Monto Total Vigente (PEN)', 'Monto del Donante (PEN)',
    'Monto de la Contrapartida (PEN)', 'Monto Sinergias y Otros'
]].copy()

for m_col in ['Monto Total Vigente (PEN)', 'Monto del Donante (PEN)', 'Monto de la Contrapartida (PEN)', 'Monto Sinergias y Otros']:
    fact_subdonacion[m_col] = pd.to_numeric(fact_subdonacion[m_col], errors='coerce')

# Dim_Subdonacion: solo textos descriptivos (conserva IdProyecto como FK hacia Dim_Proyectos)
dim_subdonacion = dim_subdonacion.drop(
    columns=[
        # Métricas → ya en Fact_Subdonacion
        'Monto Total Vigente (PEN)', 'Monto del Donante (PEN)',
        'Monto de la Contrapartida (PEN)', 'Monto Sinergias y Otros',
        'NivelSubdonacion', 'IdSubdonacionPadre',
        'CodigoSubNivel1',
        # Ubicaciones → ya en Bridge_Subdonacion_Ubicacion
        'Ubicación / Distritos (identificador)', 'Ubicación / Distritos (alias)',
        'Localidad / Comunidades (identificador)', 'Localidad / Comunidades (alias)',
        # Documentos o enlaces sin análisis
        'Documento del Convenio', 'Adendas del Convenio', 'Documentos Adicionales al Convenio',
        'Subdonación Origen (alias)', 'Por modificar: Monto del donante'
    ], errors='ignore'
)

# =========================================================
# 10. EXPORTAR
# =========================================================

with pd.ExcelWriter(r'../data/Output/modelo_estrella_proyectos.xlsx', engine="openpyxl") as writer:
    # Dimensiones
    dim_proyectos.to_excel(writer, sheet_name="Dim_Proyectos", index=False)
    dim_ambito.to_excel(writer, sheet_name="Dim_Ambito", index=False)
    dim_ubicacion_proyecto.to_excel(writer, sheet_name="Dim_Ubicacion_Proyecto", index=False)
    dim_ubicacion_subdonacion.to_excel(writer, sheet_name="Dim_Ubicacion_Subdonacion", index=False)
    dim_subdonacion.to_excel(writer, sheet_name="Dim_Subdonacion", index=False)

    # Hechos
    fact_proyectos.to_excel(writer, sheet_name="Fact_Proyectos", index=False)
    fact_informes.to_excel(writer, sheet_name="Fact_Informes", index=False)
    fact_subdonacion.to_excel(writer, sheet_name="Fact_Subdonacion", index=False)

    # Bridges
    bridge_proyecto_ubicacion.to_excel(writer, sheet_name="Bridge_Proyecto_Ubicacion", index=False)
    bridge_proyecto_ambito.to_excel(writer, sheet_name="Bridge_Proyecto_Ambito", index=False)
    bridge_subdonacion_ubicacion.to_excel(writer, sheet_name="Bridge_Subdonacion_Ubicacion", index=False)

print("Modelo estrella Kimball v7 generado correctamente")
