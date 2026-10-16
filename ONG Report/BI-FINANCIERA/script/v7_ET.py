# -*- coding: utf-8 -*-
"""
POST-v7_ET.py - Modelo estrella de finanzas v3 (Evolución de POST-v6_ET.py)

Novedades en v7:
  1. Estructura de carpetas: data/Input (entrada) y data/Output (salida).
  2. Fuente maestra en Excel: Se eliminan diccionarios hardcodeados (map_pog,
     map_responsabilidad, map_area_tematica). Los metadatos se leen directamente de
     'POST-Profonanpe - entidades.xlsx':
       - Tipo de Intervención
       - Área Temática
       - Rol de PROFONANPE
       - Seguimiento
       - Tipo Recurso Global (reemplaza a TipoPOG)
       - Moneda
       - Monto del Donante (con dolarización)
       - Estado Plan Operativo Anual (determina EstadoProyecto)
  3. Dolarización:
       - USD: * 1.0
       - EUR: * 1.1
       - CAD: * 0.72
       - PEN: / 3.4
  4. Proyecto 045 (Desagregación por Donante):
       Los donantes de 045 (ej. 045034, 045082, 045090, 045092) se tratan como
       proyectos independientes de 6 dígitos tanto en DimProyecto como en FactFinanciero.
  5. Estado del Proyecto:
       Determinado por 'Estado Plan Operativo Anual':
       'Sí, con ejecución realizada o programada' -> 'En ejecución'
       'No, sin ejecución' o no registrado       -> 'Cerrado'
  6. FactPOG:
       - 'Convenio': Monto fijo en USD obtenido de 'Monto del Donante' en entidades.
       - 'Ejecución Acumulada': Cálculo dinámico anual de suma acumulada (cumsum).
"""

import os
import re
import glob
import unicodedata
import datetime
import numpy as np
import pandas as pd

# =========================================================
# 0. CONFIGURACIÓN DE RUTAS Y ENTORNO
# =========================================================

DIRECTORIO_ACTUAL = os.path.dirname(os.path.abspath(__file__))

# Detección flexible de la carpeta 'data': busca en el directorio actual o en el padre
if os.path.exists(os.path.join(DIRECTORIO_ACTUAL, "data", "Input")):
    BASE_DIR = DIRECTORIO_ACTUAL
elif os.path.exists(os.path.join(os.path.dirname(DIRECTORIO_ACTUAL), "data", "Input")):
    BASE_DIR = os.path.dirname(DIRECTORIO_ACTUAL)
else:
    BASE_DIR = os.path.dirname(DIRECTORIO_ACTUAL)

CARPETA_INPUT = os.path.join(BASE_DIR, "data", "Input")
CARPETA_OUTPUT = os.path.join(BASE_DIR, "data", "Output")
ARCHIVO_SALIDA = os.path.join(CARPETA_OUTPUT, "modelo_estrella_finanzas.xlsx")

os.makedirs(CARPETA_INPUT, exist_ok=True)
os.makedirs(CARPETA_OUTPUT, exist_ok=True)

NIVEL_POR_LONGITUD = {3: 1, 6: 2, 9: 3, 13: 4, 15: 5}
MESES = [f"FN_IMP_EJEC_ME{m:02d}" for m in range(1, 13)]
NUMERICAS = (
    ["FN_IMP_PRAN"] + MESES +
    ["FN_IMP_TOTA_EJEC", "FN_IMP_COAN", "FN_IMP_TOTA", "FN_IMP_SALD"]
)
COLUMNAS_REQUERIDAS_FINANZAS = (
    ["FI_NUM_ANNO", "FS_COD_PART", "FS_COD_PART_CORT", "FS_DES_PART"] + NUMERICAS
)
TOLERANCIA = 0.01
ADVERTENCIAS = []


def avisar(msg):
    ADVERTENCIAS.append(msg)
    print(f"  [AVISO] {msg}")


# =========================================================
# 1. UTILIDADES Y NORMALIZACIÓN LINGÜÍSTICA
# =========================================================

def quitar_tildes(texto):
    """Quita tildes conservando la Ñ."""
    if pd.isna(texto):
        return texto
    t = str(texto).replace("Ñ", "\u0001").replace("ñ", "\u0002")
    t = "".join(
        c for c in unicodedata.normalize("NFD", t)
        if unicodedata.category(c) != "Mn"
    )
    return t.replace("\u0001", "Ñ").replace("\u0002", "ñ")


def n_tildes(texto):
    return sum(a != b for a, b in zip(texto, quitar_tildes(texto)))


def limpiar_nombre(texto):
    """Espacios dobles, comillas sueltas en los bordes, signos; MAYÚSCULAS."""
    if pd.isna(texto):
        return ""
    t = unicodedata.normalize("NFC", str(texto)).replace("\xa0", " ")
    t = t.strip().strip("'\"`´").strip()
    t = re.sub(r"\s+", " ", t)
    t = re.sub(r"\s+([:;,.)])", r"\1", t)
    t = re.sub(r"\(\s+", "(", t)
    return t.upper()


def construir_mapa_tildes(nombres_limpios):
    """Si dos escrituras solo difieren en tildes, gana la que tiene tildes."""
    frecuencia = nombres_limpios.value_counts()
    mapa = {}
    for clave, grupo in pd.Series(frecuencia.index).groupby(
        pd.Series(frecuencia.index).map(quitar_tildes)
    ):
        mejor = max(grupo, key=lambda v: (n_tildes(v), frecuencia[v]))
        for v in grupo:
            mapa[v] = mejor
    return mapa


def normalizar_codigo_entidades(codigo):
    """
    Normaliza el código del Excel de entidades:
    - Si tiene <= 3 dígitos -> zfill(3) (ej. '15' -> '015', '132' -> '132')
    - Si tiene 4 a 6 dígitos -> zfill(6) (ej. '45034' -> '045034', donantes de 045)
    """
    if pd.isna(codigo):
        return ""
    s = str(codigo).strip()
    if re.fullmatch(r"\d+\.0+", s):
        s = s.split(".")[0]
    s = s.split("-")[0].strip()
    if s.isdigit():
        if len(s) <= 3:
            return s.zfill(3)
        elif len(s) <= 6:
            return s.zfill(6)
    return s


def normalizar_cod_part(x):
    """
    Código como texto para finanzas. Limpia sufijos y repone ceros si Excel los perdió.
    Ejemplos:
      '15' -> '015'
      '45' -> '045'
      '045034 -045.2' -> '045034'
      '45007' -> '045007'
      '015006001SC0100' -> '015006001SC0100'
    """
    if pd.isna(x):
        return ""
    s = str(x).strip().upper()
    if "-" in s:
        part0 = s.split("-")[0].strip()
        if part0.isdigit():
            s = part0
    s = re.sub(r"\s+", "", s)
    if re.fullmatch(r"\d+\.0+", s):
        s = s.split(".")[0]
    if s.isdigit():
        if len(s) <= 3:
            s = s.zfill(3)
        elif len(s) == 5:
            s = s.zfill(6)
        elif len(s) == 8:
            s = s.zfill(9)
    return s


def dolarizar_monto(monto, moneda):
    """
    Convierte montos a USD según moneda:
      - USD: Monto * 1.0
      - EUR: Monto * 1.1
      - CAD: Monto * 0.72
      - PEN: Monto / 3.4
    """
    if pd.isna(monto) or monto == "":
        return 0.0
    try:
        val = float(monto)
    except Exception:
        return 0.0

    moneda_str = str(moneda).upper()
    if "EUR" in moneda_str or "EURO" in moneda_str:
        return round(val * 1.1, 2)
    elif "CAD" in moneda_str or "CANADIENSE" in moneda_str:
        return round(val * 0.72, 2)
    elif "PEN" in moneda_str or "SOL" in moneda_str:
        return round(val / 3.4, 2)
    else:  # Por defecto USD u otra moneda dolarizada
        return round(val * 1.0, 2)


# =========================================================
# 2. LECTURA DE ENTIDADES (FUENTE MAESTRA DE NEGOCIO)
# =========================================================

print("\n--- PASO 1: LECTURA DE ENTIDADES ---")
archivos_entidades = [
    f for f in glob.glob(os.path.join(CARPETA_INPUT, "*.xlsx"))
    if "entidades" in os.path.basename(f).lower() and "~$" not in os.path.basename(f)
]

df_entidades_maestro = pd.DataFrame()
map_departamentos = {
    'AMAZONAS': '01', 'ANCASH': '02', 'APURIMAC': '03', 'AREQUIPA': '04',
    'AYACUCHO': '05', 'CAJAMARCA': '06', 'CALLAO': '07', 'CUSCO': '08',
    'HUANCAVELICA': '09', 'HUANUCO': '10', 'ICA': '11', 'JUNIN': '12',
    'LA LIBERTAD': '13', 'LAMBAYEQUE': '14', 'LIMA': '15', 'LORETO': '16',
    'MADRE DE DIOS': '17', 'MOQUEGUA': '18', 'PASCO': '19', 'PIURA': '20',
    'PUNO': '21', 'SAN MARTIN': '22', 'TACNA': '23', 'TUMBES': '24', 'UCAYALI': '25'
}

df_ent_ubicacion = pd.DataFrame()

if not archivos_entidades:
    avisar(f"No se encontró archivo de entidades en '{CARPETA_INPUT}'. Se usará metadatos por defecto.")
else:
    archivo_ent = archivos_entidades[0]
    print(f"  Leyendo entidades desde: {os.path.basename(archivo_ent)}")
    ent_raw = pd.read_excel(archivo_ent)
    ent_raw.columns = ent_raw.columns.astype(str).str.strip()

    # Identificación flexible de columnas requeridas
    def buscar_col(patron, cols):
        for c in cols:
            if re.search(patron, c, re.IGNORECASE):
                return c
        return None

    c_cod = buscar_col(r"^c[oó]digo", ent_raw.columns) or ent_raw.columns[0]
    c_interv = buscar_col(r"tipo.*intervenci[oó]n", ent_raw.columns)
    c_area = buscar_col(r"[aá]rea.*tem[aá]tica", ent_raw.columns)
    c_rol = buscar_col(r"rol.*profonanpe", ent_raw.columns)
    c_seg = buscar_col(r"seguimiento", ent_raw.columns)
    c_trg = buscar_col(r"tipo.*recurso.*global", ent_raw.columns)
    c_mon = buscar_col(r"moneda", ent_raw.columns)
    c_monto = buscar_col(r"monto.*donante", ent_raw.columns)
    c_poa = buscar_col(r"estado.*plan.*operativo", ent_raw.columns)
    c_depto = buscar_col(r"departamento", ent_raw.columns)

    ent = pd.DataFrame()
    ent["CodigoProyecto"] = ent_raw[c_cod].apply(normalizar_codigo_entidades)
    ent = ent[ent["CodigoProyecto"].ne("")].copy()

    ent["TipoIntervencion"] = ent_raw[c_interv].fillna("Sin Clasificar").astype(str).str.strip() if c_interv else "Sin Clasificar"
    ent["AreaTematica"] = ent_raw[c_area].fillna("Sin Clasificar").astype(str).str.strip() if c_area else "Sin Clasificar"
    ent["RolProfonanpe"] = ent_raw[c_rol].fillna("Sin Clasificar").astype(str).str.strip() if c_rol else "Sin Clasificar"
    ent["Seguimiento"] = ent_raw[c_seg].fillna("Sin Clasificar").astype(str).str.strip() if c_seg else "Sin Clasificar"
    ent["TipoRecursoGlobal"] = ent_raw[c_trg].fillna("Sin Clasificar").astype(str).str.strip() if c_trg else "Sin Clasificar"
    ent["Moneda"] = ent_raw[c_mon].fillna("No Registrada").astype(str).str.strip() if c_mon else "No Registrada"

    # Montos y Dolarización
    ent["MontoDonanteOriginal"] = pd.to_numeric(ent_raw[c_monto], errors="coerce").fillna(0.0) if c_monto else 0.0
    ent["MontoDonanteUSD"] = [
        dolarizar_monto(m, mon) for m, mon in zip(ent["MontoDonanteOriginal"], ent["Moneda"])
    ]

    # Estado Plan Operativo Anual (con sus 3 valores posibles en el modelo)
    ent["EstadoPlanOperativoAnual"] = ent_raw[c_poa].fillna("No Registrado").astype(str).str.strip() if c_poa else "No Registrado"

    df_entidades_maestro = ent.drop_duplicates(subset=["CodigoProyecto"]).reset_index(drop=True)
    print(f"  Entidades procesadas: {len(df_entidades_maestro)} proyectos/donantes únicos.")

    # Guardar copia para ubicación
    if c_depto:
        df_ent_ubicacion = pd.DataFrame({
            "CodigoProyecto": ent_raw[c_cod].apply(normalizar_codigo_entidades),
            "Departamentos": ent_raw[c_depto].astype(str)
        })


# =========================================================
# 3. LECTURA DE FINANZAS
# =========================================================

print("\n--- PASO 2: LECTURA DE FINANZAS (INPUT) ---")
archivos_finanzas = sorted(
    f for ext in ("xlsb", "xlsx")
    for f in glob.glob(os.path.join(CARPETA_INPUT, f"*.{ext}"))
    if "~$" not in os.path.basename(f) and "entidades" not in os.path.basename(f).lower()
)

if not archivos_finanzas:
    raise SystemExit(f"Error: No hay archivos financieros (.xlsb / .xlsx) en '{CARPETA_INPUT}'")


def leer_archivo_finanzas(ruta):
    dtypes = {"FS_COD_PART": str, "FS_COD_PART_CORT": str}
    if ruta.lower().endswith(".xlsb"):
        d = pd.read_excel(ruta, engine="pyxlsb", dtype=dtypes)
    else:
        d = pd.read_excel(ruta, dtype=dtypes)
    d.columns = d.columns.astype(str).str.strip()
    d["ArchivoOrigen"] = os.path.basename(ruta)
    return d


lista_fin = []
for archivo in archivos_finanzas:
    print(f"  Leyendo: {os.path.basename(archivo)}")
    df_temp = leer_archivo_finanzas(archivo)
    faltan = [c for c in COLUMNAS_REQUERIDAS_FINANZAS if c not in df_temp.columns]
    if faltan:
        raise SystemExit(f"Error en {archivo}: faltan columnas requeridas {faltan}")
    lista_fin.append(df_temp)

df_fin = pd.concat(lista_fin, ignore_index=True)


# =========================================================
# 4. LIMPIEZA Y CLASIFICACIÓN JERÁRQUICA
# =========================================================

df_fin["Anio"] = pd.to_numeric(df_fin["FI_NUM_ANNO"], errors="coerce")
df_fin = df_fin[df_fin["Anio"].notna()].copy()
df_fin["Anio"] = df_fin["Anio"].astype(int)

df_fin["Cod"] = df_fin["FS_COD_PART"].apply(normalizar_cod_part)
df_fin["Longitud"] = df_fin["Cod"].str.len()
df_fin["Nivel"] = df_fin["Longitud"].map(NIVEL_POR_LONGITUD)

for col in NUMERICAS:
    df_fin[col] = pd.to_numeric(df_fin[col], errors="coerce").fillna(0.0)

# Auditoría de códigos raros
qa_codigos_raros = df_fin[df_fin["Nivel"].isna()][
    ["ArchivoOrigen", "Anio", "FS_COD_PART", "FS_DES_PART"]
].copy()
if len(qa_codigos_raros):
    avisar(f"{len(qa_codigos_raros)} filas con código de longitud no válida (ver QA_CodigosRaros); se excluyen.")
df_fin = df_fin[df_fin["Nivel"].notna()].copy()
df_fin["Nivel"] = df_fin["Nivel"].astype(int)

# Auditoría de duplicados
dup = df_fin[df_fin.duplicated(["Anio", "Cod"], keep=False)]
qa_duplicados = dup[["ArchivoOrigen", "Anio", "Cod", "FS_DES_PART", "FN_IMP_TOTA_EJEC"]].sort_values(["Anio", "Cod"])
if len(qa_duplicados):
    avisar(f"{len(qa_duplicados)} filas con código repetido en el mismo año (ver QA_Duplicados).")

# Nombres limpios con unificación de tildes
df_fin["Nombre_Original"] = df_fin["FS_DES_PART"].astype(str).str.strip()
df_fin["_base"] = df_fin["Nombre_Original"].apply(limpiar_nombre)
mapa_tildes = construir_mapa_tildes(df_fin["_base"])
df_fin["Nombre"] = df_fin["_base"].map(mapa_tildes)

anios = sorted(int(a) for a in df_fin["Anio"].unique())
print(f"  Años cargados: {anios} | Total registros por nivel: {df_fin.groupby('Nivel').size().to_dict()}")


# =========================================================
# 5. CONSTRUCCIÓN DE DIMENSIONES (CON DESAGREGACIÓN DE 045)
# =========================================================

def ultimo_nombre(d, cod_col="Cod"):
    """Devuelve un registro único por código usando el nombre del año más reciente."""
    return (d.sort_values("Anio", ascending=False)
             .drop_duplicates(cod_col)
             .sort_values(cod_col)
             .reset_index(drop=True))


# ---- DimProyecto (Tratamiento especial para 045)
n1 = df_fin[df_fin["Nivel"] == 1].copy()
n2 = df_fin[df_fin["Nivel"] == 2].copy()
n5 = df_fin[df_fin["Nivel"] == 5].copy()

# A. Proyectos regulares (Nivel 1 que NO son 045)
n1_regulares = n1[n1["Cod"] != "045"].copy()
p_regulares = ultimo_nombre(n1_regulares)[["Cod", "Nombre", "Nombre_Original"]].rename(columns={
    "Cod": "CodigoProyecto", "Nombre": "Proyecto", "Nombre_Original": "Proyecto_Original"
})

# B. Proyectos desagregados de 045 (Nivel 2 o 5 que comienzan con 045)
# Cada donante de 045 es un proyecto independiente con código de 6 dígitos
d45_en_n2 = n2[n2["Cod"].str.startswith("045")].copy()
p45 = pd.DataFrame()
if len(d45_en_n2):
    p45 = ultimo_nombre(d45_en_n2)[["Cod", "Nombre", "Nombre_Original"]].rename(columns={
        "Cod": "CodigoProyecto", "Nombre": "Proyecto", "Nombre_Original": "Proyecto_Original"
    })
else:
    # Si no vinieron filas de nivel 2 para 045, buscamos códigos en nivel 5
    d45_en_n5 = n5[n5["Cod"].str.startswith("045")].copy()
    if len(d45_en_n5):
        d45_en_n5["Cod6"] = d45_en_n5["Cod"].str[:6]
        p45 = ultimo_nombre(d45_en_n5, "Cod6")[["Cod6", "Nombre", "Nombre_Original"]].rename(columns={
            "Cod6": "CodigoProyecto", "Nombre": "Proyecto", "Nombre_Original": "Proyecto_Original"
        })

dim_proyecto = pd.concat([p_regulares, p45], ignore_index=True).sort_values("CodigoProyecto").reset_index(drop=True)

# Cruzar con información de Entidades (Fuente Maestra)
if not df_entidades_maestro.empty:
    dim_proyecto = dim_proyecto.merge(
        df_entidades_maestro,
        on="CodigoProyecto",
        how="left"
    )
    dim_proyecto["TipoIntervencion"] = dim_proyecto["TipoIntervencion"].fillna("Sin Clasificar")
    dim_proyecto["AreaTematica"] = dim_proyecto["AreaTematica"].fillna("Sin Clasificar")
    dim_proyecto["RolProfonanpe"] = dim_proyecto["RolProfonanpe"].fillna("Sin Clasificar")
    dim_proyecto["Seguimiento"] = dim_proyecto["Seguimiento"].fillna("Sin Clasificar")
    dim_proyecto["TipoRecursoGlobal"] = dim_proyecto["TipoRecursoGlobal"].fillna("Sin Clasificar")
    dim_proyecto["Moneda"] = dim_proyecto["Moneda"].fillna("No Registrada")
    dim_proyecto["MontoDonanteOriginal"] = dim_proyecto["MontoDonanteOriginal"].fillna(0.0)
    dim_proyecto["MontoDonanteUSD"] = dim_proyecto["MontoDonanteUSD"].fillna(0.0)
    dim_proyecto["EstadoPlanOperativoAnual"] = dim_proyecto["EstadoPlanOperativoAnual"].fillna("No Registrado")
else:
    dim_proyecto["TipoIntervencion"] = "Sin Clasificar"
    dim_proyecto["AreaTematica"] = "Sin Clasificar"
    dim_proyecto["RolProfonanpe"] = "Sin Clasificar"
    dim_proyecto["Seguimiento"] = "Sin Clasificar"
    dim_proyecto["TipoRecursoGlobal"] = "Sin Clasificar"
    dim_proyecto["Moneda"] = "No Registrada"
    dim_proyecto["MontoDonanteOriginal"] = 0.0
    dim_proyecto["MontoDonanteUSD"] = 0.0
    dim_proyecto["EstadoPlanOperativoAnual"] = "No Registrado"

# Alerta sin movimiento (Año más reciente en los datos)
anio_reciente = max(anios)
if len(d45_en_n2) or len(n5[n5["Cod"].str.startswith("045")]):
    # Proyectos con movimiento en año reciente (regulares N1 + donantes 045 en N5/N2)
    reg_mov = n1[(n1["Anio"] == anio_reciente) & (n1["Cod"] != "045") & (n1["FN_IMP_TOTA_EJEC"].abs() > 0)]["Cod"]
    d45_mov = n5[(n5["Anio"] == anio_reciente) & (n5["Cod"].str.startswith("045")) & (n5["FN_IMP_TOTA_EJEC"].abs() > 0)]["Cod"].str[:6]
    proy_con_mov = set(reg_mov).union(set(d45_mov))
else:
    proy_con_mov = set(n1[(n1["Anio"] == anio_reciente) & (n1["FN_IMP_TOTA_EJEC"].abs() > 0)]["Cod"])

dim_proyecto["AlertaSinMovimiento"] = np.where(
    (dim_proyecto["EstadoPlanOperativoAnual"].str.contains(r"^S[ií]|con ejecuci[oó]n", case=False, na=False)) & (~dim_proyecto["CodigoProyecto"].isin(proy_con_mov)),
    "Sí",
    "No"
)

dim_proyecto.insert(0, "IdProyecto", range(1, len(dim_proyecto) + 1))

# Reordenar columnas de DimProyecto (sin EstadoProyecto)
cols_dim_proyecto = [
    "IdProyecto", "CodigoProyecto", "Proyecto", "Proyecto_Original",
    "EstadoPlanOperativoAnual", "AlertaSinMovimiento",
    "TipoIntervencion", "AreaTematica", "RolProfonanpe", "Seguimiento",
    "TipoRecursoGlobal", "Moneda", "MontoDonanteOriginal", "MontoDonanteUSD"
]
dim_proyecto = dim_proyecto[cols_dim_proyecto]

# ---- Donante / Ejecutor / Unidad Ejecutora / Componente
def dim_segmento(nivel, ini, fin, col_cod, col_nom, col_id):
    d = df_fin[df_fin["Nivel"] == nivel].copy()
    d["Seg"] = d["Cod"].str[ini:fin]
    u = ultimo_nombre(d, "Seg")[["Seg", "Nombre", "Nombre_Original"]]
    u.columns = [col_cod, col_nom, col_nom + "_Original"]
    u.insert(0, col_id, range(1, len(u) + 1))
    return u

dim_donante = dim_segmento(2, 3, 6, "CodigoDonante", "Donante", "IdDonante")
dim_ejecutor = dim_segmento(3, 6, 9, "CodigoEjecutor", "Ejecutor", "IdEjecutor")
dim_unidad = dim_segmento(4, 9, 13, "CodigoUnidadEjecutora", "UnidadEjecutora", "IdUnidadEjecutora")

dim_comp = ultimo_nombre(n5)
c_comp = dim_comp["Cod"]
dim_comp = pd.DataFrame({
    "CodigoComponenteCompleto": c_comp,
    "CodigoComponente": c_comp.str[13:15],
    "Componente": dim_comp["Nombre"],
    "Componente_Original": dim_comp["Nombre_Original"],
})
dim_comp.insert(0, "IdComponente", range(1, len(dim_comp) + 1))

# DimFecha
dim_fecha = pd.DataFrame({"Anio": anios})
dim_fecha["Fecha"] = pd.to_datetime(dict(year=dim_fecha["Anio"], month=1, day=1))
dim_fecha.insert(0, "IdFecha", range(1, len(dim_fecha) + 1))
dim_fecha = dim_fecha[["IdFecha", "Fecha", "Anio"]]


# =========================================================
# 6. CONSTRUCCIÓN DE FACT FINANCIERO (NIVEL 5)
# =========================================================

print("\n--- PASO 3: CONSTRUCCIÓN DE FACT FINANCIERO ---")
fact = n5.copy()

# Asignación de CodigoProyecto: 6 dígitos para 045 (donante propio), 3 para el resto
fact["CodigoProyecto"] = np.where(
    fact["Cod"].str.startswith("045"),
    fact["Cod"].str[:6],
    fact["Cod"].str[:3]
)
fact["CodigoDonante"] = fact["Cod"].str[3:6]
fact["CodigoEjecutor"] = fact["Cod"].str[6:9]
fact["CodigoUnidadEjecutora"] = fact["Cod"].str[9:13]

fact["IdProyecto"] = fact["CodigoProyecto"].map(dict(zip(dim_proyecto.CodigoProyecto, dim_proyecto.IdProyecto)))
fact["IdDonante"] = fact["CodigoDonante"].map(dict(zip(dim_donante.CodigoDonante, dim_donante.IdDonante)))
fact["IdEjecutor"] = fact["CodigoEjecutor"].map(dict(zip(dim_ejecutor.CodigoEjecutor, dim_ejecutor.IdEjecutor)))
fact["IdUnidadEjecutora"] = fact["CodigoUnidadEjecutora"].map(dict(zip(dim_unidad.CodigoUnidadEjecutora, dim_unidad.IdUnidadEjecutora)))
fact["IdComponente"] = fact["Cod"].map(dict(zip(dim_comp.CodigoComponenteCompleto, dim_comp.IdComponente)))
fact["IdFecha"] = fact["Anio"].map(dict(zip(dim_fecha.Anio, dim_fecha.IdFecha)))

renombres_fact = {
    "FN_IMP_PRAN": "POA",
    "FN_IMP_TOTA_EJEC": "EjecutadoAnual",
    "FN_IMP_COAN": "Comprometido",
    "FN_IMP_TOTA": "Total",
    "FN_IMP_SALD": "Saldo",
}
renombres_fact.update({f"FN_IMP_EJEC_ME{m:02d}": f"Ejec_M{m:02d}" for m in range(1, 13)})

fact_finanzas = fact.sort_values(["Anio", "IdComponente"]).rename(columns=renombres_fact)[
    ["IdFecha", "IdProyecto", "IdDonante", "IdEjecutor",
     "IdUnidadEjecutora", "IdComponente", "POA"]
    + [f"Ejec_M{m:02d}" for m in range(1, 13)]
    + ["EjecutadoAnual", "Comprometido", "Total", "Saldo"]
].reset_index(drop=True)


# =========================================================
# 7. CONSTRUCCIÓN DE FACT POG (CONVENIO VS EJECUCIÓN ACUMULADA)
# =========================================================

print("\n--- PASO 4: CONSTRUCCIÓN DE FACT POG ---")
# Agrupar ejecución por Proyecto y Año desde FactFinanciero (respeta la desagregación de 045)
pog = fact.groupby(["CodigoProyecto", "Anio"], as_index=False)["FN_IMP_TOTA_EJEC"].sum()
pog = pog.sort_values(["CodigoProyecto", "Anio"]).reset_index(drop=True)

# Cruzar con TipoRecursoGlobal y MontoDonanteUSD de DimProyecto
pog = pog.merge(
    dim_proyecto[["CodigoProyecto", "IdProyecto", "TipoRecursoGlobal", "MontoDonanteUSD"]],
    on="CodigoProyecto",
    how="left"
)

# Cálculo según TipoRecursoGlobal
pog["POG"] = np.nan

# 1. Convenio -> Monto fijo en USD de entidades
mask_convenio = pog["TipoRecursoGlobal"].astype(str).str.contains("Convenio", case=False, na=False)
pog.loc[mask_convenio, "POG"] = pog.loc[mask_convenio, "MontoDonanteUSD"]

# 2. Ejecución Acumulada -> Suma acumulada de la ejecución histórica (cumsum)
mask_acum = pog["TipoRecursoGlobal"].astype(str).str.contains("Acumulad", case=False, na=False)
pog.loc[mask_acum, "POG"] = pog.loc[mask_acum].groupby("CodigoProyecto")["FN_IMP_TOTA_EJEC"].cumsum()

# 3. Resto (Sin Clasificar o no registrado en entidades) -> 0.0
# Un POG en 0 indica que falta registrar el proyecto en entidades (reporte vtrack)
mask_resto = pog["POG"].isna()
pog.loc[mask_resto, "POG"] = 0.0

pog["IdFecha"] = pog["Anio"].map(dict(zip(dim_fecha.Anio, dim_fecha.IdFecha)))
fact_pog = pog[["IdFecha", "IdProyecto", "TipoRecursoGlobal", "POG"]].sort_values(["IdFecha", "IdProyecto"]).reset_index(drop=True)


# =========================================================
# 8. DIMENSIÓN UBICACIÓN Y BRIDGE
# =========================================================

print("\n--- PASO 5: PROCESAMIENTO DE UBICACIONES ---")
dim_ubicacion = pd.DataFrame(columns=["IdUbicacion", "Departamento", "Ubigeo"])
bridge_proyecto_ubicacion = pd.DataFrame(columns=["IdProyecto", "IdUbicacion"])
qa_sin_ubicacion = dim_proyecto[["IdProyecto", "CodigoProyecto", "Proyecto"]]
qa_depto_sin_ubigeo = pd.DataFrame(columns=["Departamento"])

if not df_ent_ubicacion.empty:
    df_ub = df_ent_ubicacion.copy()
    df_ub["Departamento"] = df_ub["Departamentos"].astype(str).str.upper().str.split(";")
    df_ub = df_ub.explode("Departamento")
    df_ub["Departamento"] = df_ub["Departamento"].str.strip().apply(quitar_tildes)
    df_ub = df_ub[df_ub["Departamento"].notna() & df_ub["Departamento"].ne("") & df_ub["Departamento"].ne("NAN")]

    dim_ubicacion = df_ub[["Departamento"]].drop_duplicates().reset_index(drop=True)
    dim_ubicacion["Ubigeo"] = dim_ubicacion["Departamento"].map(map_departamentos)
    qa_depto_sin_ubigeo = dim_ubicacion[dim_ubicacion["Ubigeo"].isna()][["Departamento"]]
    dim_ubicacion = dim_ubicacion.dropna(subset=["Ubigeo"]).sort_values("Ubigeo").reset_index(drop=True)
    dim_ubicacion.insert(0, "IdUbicacion", range(1, len(dim_ubicacion) + 1))

    # Bridge
    br = df_ub[["CodigoProyecto", "Departamento"]].merge(
        dim_proyecto[["CodigoProyecto", "IdProyecto"]],
        on="CodigoProyecto",
        how="inner"
    )
    br["Ubigeo"] = br["Departamento"].map(map_departamentos)
    br = br.merge(dim_ubicacion[["Ubigeo", "IdUbicacion"]], on="Ubigeo", how="inner")

    bridge_proyecto_ubicacion = (
        br[["IdProyecto", "IdUbicacion"]]
        .drop_duplicates()
        .sort_values(["IdProyecto", "IdUbicacion"])
        .reset_index(drop=True)
    )

    con_ubic = set(bridge_proyecto_ubicacion["IdProyecto"])
    qa_sin_ubicacion = dim_proyecto[~dim_proyecto["IdProyecto"].isin(con_ubic)][["IdProyecto", "CodigoProyecto", "Proyecto"]]
    if len(qa_depto_sin_ubigeo):
        avisar(f"{len(qa_depto_sin_ubigeo)} departamentos de entidades sin ubigeo asignado (ver QA_DepartamentosSinUbigeo).")


# =========================================================
# 9. QA: AUDITORÍAS Y CONTROL DE CUADRE
# =========================================================

print("\n--- PASO 6: EJECUTANDO PRUEBAS DE QA ---")
# A. Cuadre entre Nivel 1 y FactFinanciero (teniendo en cuenta que 045 en Nivel 1 cuadra contra la suma de sus 045XXX)
medidas = [("POA", "FN_IMP_PRAN", "POA"), ("Ejec", "FN_IMP_TOTA_EJEC", "EjecutadoAnual"),
           ("Comprom", "FN_IMP_COAN", "Comprometido"), ("Saldo", "FN_IMP_SALD", "Saldo")]

a_n1 = n1.groupby(["Anio", "Cod"], as_index=False)[[m[1] for m in medidas]].sum()
b_fact = fact.copy()
# Para cuadrar contra N1, el código de comparación de 045 vuelve a ser 045
b_fact["CodComparacion"] = np.where(b_fact["Cod"].str.startswith("045"), "045", b_fact["Cod"].str[:3])
b_fact = b_fact.groupby(["Anio", "CodComparacion"], as_index=False)[[m[1] for m in medidas]].sum()
b_fact = b_fact.rename(columns={"CodComparacion": "Cod"})

cuadre = a_n1.merge(b_fact, on=["Anio", "Cod"], how="outer", suffixes=("_N1", "_Fact")).fillna(0.0)
cuadre = cuadre.merge(p_regulares[["CodigoProyecto", "Proyecto"]], left_on="Cod", right_on="CodigoProyecto", how="left")
qa_cuadre = cuadre[["Anio", "Cod", "Proyecto"]].rename(columns={"Cod": "CodigoProyecto"})

difs = []
for etiqueta, col, _ in medidas:
    qa_cuadre[f"{etiqueta}_Nivel1"] = cuadre[col + "_N1"].values
    qa_cuadre[f"{etiqueta}_Fact"] = cuadre[col + "_Fact"].values
    qa_cuadre[f"Dif_{etiqueta}"] = (cuadre[col + "_Fact"] - cuadre[col + "_N1"]).round(2).values
    difs.append(f"Dif_{etiqueta}")

qa_cuadre["Resultado"] = np.where(qa_cuadre[difs].abs().max(axis=1) > TOLERANCIA, "REVISAR", "OK")
qa_cuadre = qa_cuadre.sort_values(["Anio", "CodigoProyecto"]).reset_index(drop=True)

n_rev = (qa_cuadre["Resultado"] == "REVISAR").sum()
if n_rev:
    avisar(f"{n_rev} proyecto-año NO cuadran entre el nivel proyecto y la suma de componentes (ver QA_Cuadre).")

# B. Auditoría de nombres limpios
qa_nombres_limpieza = (
    df_fin[df_fin["Nombre_Original"] != df_fin["Nombre"]]
    [["Nivel", "Cod", "Nombre_Original", "Nombre"]]
    .drop_duplicates("Nombre_Original")
    .rename(columns={"Cod": "Codigo", "Nombre": "Nombre_Limpio"})
    .sort_values(["Nivel", "Codigo"])
)

# C. Renombres entre años
x_ren = df_fin.groupby(["Nivel", "Cod"])["Nombre"].agg(lambda v: sorted(set(v)))
x_ren = x_ren[x_ren.map(len) > 1].reset_index()
x_ren["Nombres"] = x_ren["Nombre"].map(lambda v: " | ".join(v))
qa_renombres = x_ren[["Nivel", "Cod", "Nombres"]].rename(columns={"Cod": "Codigo"})

# D. Segmentos en componentes sin fila propia en su nivel
codigos_comp = dim_comp["CodigoComponenteCompleto"]
qa_segmentos = []
for col, dim, nivel, ini, fin in [("CodigoProyecto", dim_proyecto, "Proyecto", 0, 3),
                                  ("CodigoDonante", dim_donante, "Donante", 3, 6),
                                  ("CodigoEjecutor", dim_ejecutor, "Ejecutor", 6, 9),
                                  ("CodigoUnidadEjecutora", dim_unidad, "UnidadEjecutora", 9, 13)]:
    if nivel == "Proyecto":
        # Para proyecto, recordar que 045 usa 6 caracteres
        faltan = sorted(set(np.where(codigos_comp.str.startswith("045"), codigos_comp.str[:6], codigos_comp.str[:3])) - set(dim[col]))
    else:
        faltan = sorted(set(codigos_comp.str[ini:fin]) - set(dim[col]))
    for f_ in faltan:
        qa_segmentos.append({"Nivel": nivel, "Codigo": f_})
qa_segmentos = pd.DataFrame(qa_segmentos, columns=["Nivel", "Codigo"])


# =========================================================
# 10. EXPORTACIÓN A EXCEL
# =========================================================

print(f"\n--- PASO 7: EXPORTANDO RESULTADOS A '{ARCHIVO_SALIDA}' ---")
hojas = [
    ("FactFinanciero", fact_finanzas),
    ("FactPOG", fact_pog),
    ("DimProyecto", dim_proyecto),
    ("DimDonante", dim_donante),
    ("DimEjecutor", dim_ejecutor),
    ("DimUnidadEjecutora", dim_unidad),
    ("DimComponente", dim_comp),
    ("DimFecha", dim_fecha),
    ("DimUbicacion", dim_ubicacion),
    ("BridgeProyectoUbicacion", bridge_proyecto_ubicacion),
    ("QA_Cuadre", qa_cuadre),
    ("QA_NombresLimpieza", qa_nombres_limpieza),
    ("QA_RenombresEntreAnios", qa_renombres),
    ("QA_SegmentosSinFila", qa_segmentos),
    ("QA_CodigosRaros", qa_codigos_raros),
    ("QA_Duplicados", qa_duplicados),
    ("QA_ProyectosSinUbicacion", qa_sin_ubicacion),
    ("QA_DepartamentosSinUbigeo", qa_depto_sin_ubigeo),
]

with pd.ExcelWriter(ARCHIVO_SALIDA) as writer:
    for nombre, tabla in hojas:
        tabla.to_excel(writer, sheet_name=nombre, index=False)

print(f"\n[OK] Modelo estrella generado exitosamente en:\n     {ARCHIVO_SALIDA}")
print(f"  - FactFinanciero: {len(fact_finanzas)} filas")
print(f"  - DimProyecto: {len(dim_proyecto)} proyectos")
print(f"    * EstadoPlanOperativoAnual:")
for val, cnt in dim_proyecto['EstadoPlanOperativoAnual'].value_counts().items():
    print(f"      - {val}: {cnt}")
print(f"    * Con alerta sin movimiento: {(dim_proyecto.AlertaSinMovimiento == 'Sí').sum()}")
if ADVERTENCIAS:
    print(f"\n[!] Se generaron {len(ADVERTENCIAS)} aviso(s) (revisa las hojas QA_*):")
    for w in ADVERTENCIAS:
        print(f"  - {w}")
else:
    print("\n[OK] Sin avisos: cuadre perfecto.")
