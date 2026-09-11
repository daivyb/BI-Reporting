import pandas as pd
import numpy as np
import json
import os
import re
import glob
from pathlib import Path


# ===========================================================================
# CONFIGURACIÓN
# ===========================================================================
TARGET_YEAR = 2026
CURRENT_QUARTER = "TR02"  # Trimestre al cual se asignará el avance físico actual (Sciforma)

# ===========================================================================
# Utilidades reutilizadas de proceso5.py (lectura Sciforma y WBS)
# ===========================================================================

COLUMN_ALIASES = {
    "WBS": ["WBS", "EDT"],
    "ID de EDT completo": ["ID de EDT completo"],
    "ID": ["ID", "Identificador"],
    "Name": ["Name", "Nombre"],
    "Start": ["Start", "Inicio"],
    "Finish": ["Finish", "Fin"],
    "% completado": ["% completado", "% completed", "% Complete"],
    "% Comp. Prom.": ["% Comp. Prom.", "% Comp Prom", "% Comp. Prom", "% Comp Prom."],
}

REQUIRED_COLUMNS = [
    "WBS", "ID de EDT completo", "ID", "Name", "Start", "Finish", "% Comp. Prom.",
]


def normalize_text(value):
    if pd.isna(value):
        return ""
    return str(value).replace("\n", " ").replace("\r", " ").strip()


def normalize_columns(df):
    df = df.copy()
    df.columns = [normalize_text(c) for c in df.columns]
    rename_map = {}
    for canonical, aliases in COLUMN_ALIASES.items():
        alias_set = {normalize_text(a).lower() for a in aliases}
        for col in df.columns:
            if normalize_text(col).lower() in alias_set:
                rename_map[col] = canonical
                break
    df = df.rename(columns=rename_map)
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f"Faltan columnas obligatorias: {missing}")
    if "% completado" not in df.columns:
        df["% completado"] = np.nan
    return df


def find_header_row(preview):
    needed_groups = [
        {"wbs", "edt"},
        {"id de edt completo"},
        {"id", "identificador"},
        {"name", "nombre"},
        {"start", "inicio"},
        {"finish", "fin"},
        {"% comp. prom.", "% comp prom", "% comp. prom", "% comp prom."},
    ]
    for i in range(len(preview)):
        row_values = {normalize_text(v).lower() for v in preview.iloc[i].tolist()}
        hits = sum(1 for group in needed_groups if row_values & group)
        if hits >= 6:
            return i
    raise ValueError("No se pudo detectar la fila de encabezado.")


def read_sciforma_file(file_path):
    preview = pd.read_excel(file_path, header=None, nrows=15)
    header_row = find_header_row(preview)
    df = pd.read_excel(file_path, header=header_row)
    return normalize_columns(df)


def wbs_depth(wbs):
    text = normalize_text(wbs)
    if not text:
        return None
    return text.count(".")


def level_from_wbs(wbs):
    text = normalize_text(wbs)
    if text == "":
        return "Proyecto"
    depth = wbs_depth(text)
    return {0: "Componente", 1: "Subcomponente", 2: "Actividad", 3: "Tarea"}.get(depth, "")


def is_task_level(wbs):
    return wbs_depth(wbs) == 3


def parent_wbs(wbs, levels_up=1):
    parts = wbs.split(".")
    if len(parts) <= levels_up:
        return None
    return ".".join(parts[:-levels_up])


def in_target_year(start, finish, year):
    return pd.notna(start) and start.year == year and pd.notna(finish) and finish.year == year


def extract_project_code(project_name):
    text = normalize_text(project_name)
    match = re.search(r"(\d{3})", text)
    return match.group(1) if match else ""


def filter_to_valid_branches(df):
    df = df.copy()
    df["WBS"] = df["WBS"].map(normalize_text)
    task_wbs = [w for w in df.loc[df["WBS"].map(is_task_level), "WBS"].tolist() if w]
    if not task_wbs:
        return df.iloc[0:0].copy()
    valid_wbs = set(task_wbs)
    for task in task_wbs:
        for levels_up in (1, 2, 3):
            p = parent_wbs(task, levels_up)
            if p:
                valid_wbs.add(p)
    keep_mask = df["WBS"].eq("") | df["WBS"].isin(valid_wbs)
    return df[keep_mask].copy()


def recalculate_percentages(df, year):
    """Recalcula % Comp. Prom. en cascada: Tarea -> Actividad -> Subcomp -> Componente."""
    df = df.copy()
    df["WBS"] = df["WBS"].map(normalize_text)
    df["Name"] = df["Name"].map(normalize_text)
    df["__depth__"] = df["WBS"].map(wbs_depth)
    df["% Comp. Prom."] = pd.to_numeric(df["% Comp. Prom."], errors="coerce")

    tasks_all = df[df["WBS"].map(is_task_level)].copy()
    tasks_year = tasks_all[
        tasks_all.apply(lambda r: in_target_year(r["Start"], r["Finish"], year), axis=1)
    ].copy()
    if tasks_year.empty:
        return df.iloc[0:0].copy()

    tasks_year["__task_progress__"] = tasks_year["% Comp. Prom."]

    activities = (
        tasks_year.assign(__activity__=tasks_year["WBS"].map(lambda x: parent_wbs(x, 1)))
        .groupby("__activity__", dropna=True)["__task_progress__"]
        .mean()
    )
    subcomponents = (
        activities.rename_axis("WBS").reset_index(name="value")
        .assign(__subcomponent__=lambda d: d["WBS"].map(lambda x: parent_wbs(x, 1)))
        .groupby("__subcomponent__", dropna=True)["value"]
        .mean()
    )
    components = (
        subcomponents.rename_axis("WBS").reset_index(name="value")
        .assign(__component__=lambda d: d["WBS"].map(lambda x: parent_wbs(x, 1)))
        .groupby("__component__", dropna=True)["value"]
        .mean()
    )

    df.loc[df["__depth__"] == 2, "% Comp. Prom."] = df.loc[df["__depth__"] == 2, "WBS"].map(activities)
    df.loc[df["__depth__"] == 1, "% Comp. Prom."] = df.loc[df["__depth__"] == 1, "WBS"].map(subcomponents)
    df.loc[df["__depth__"] == 0, "% Comp. Prom."] = df.loc[df["__depth__"] == 0, "WBS"].map(components)

    task_mask = df["WBS"].map(is_task_level)
    task_year_mask = df.apply(lambda r: in_target_year(r["Start"], r["Finish"], year), axis=1)
    df = df[~task_mask | task_year_mask].copy()
    return filter_to_valid_branches(df)


def process_one_sciforma_file(file_path, year):
    """Procesa un archivo Sciforma y retorna el DataFrame con avance recalculado."""
    df = read_sciforma_file(file_path)
    df["Start"] = pd.to_datetime(df["Start"], errors="coerce")
    df["Finish"] = pd.to_datetime(df["Finish"], errors="coerce")

    if df.empty:
        raise ValueError(f"El archivo {file_path.name} no contiene filas de datos.")
    project_row = df.iloc[0].copy()
    project_name = normalize_text(project_row.get("Name", file_path.stem))
    df = df.iloc[1:].copy()

    df = recalculate_percentages(df, year)

    df["__depth__"] = df["WBS"].map(wbs_depth)
    df = df[df["__depth__"].notna()].copy()
    df["Proyecto"] = project_name
    df["Nivel"] = df["WBS"].map(level_from_wbs)

    # Agregar fila de proyecto (promedio de componentes)
    componentes = df[df["__depth__"] == 0].copy()
    promedio_proyecto = componentes["% Comp. Prom."].mean() if not componentes.empty else np.nan

    task_count = len(df[df["__depth__"] == 3])

    project_data = {col: np.nan for col in df.columns}
    project_data["WBS"] = ""
    project_data["Proyecto"] = project_name
    project_data["Nivel"] = "Proyecto"
    project_data["% Comp. Prom."] = promedio_proyecto
    project_data["Name"] = project_name
    project_data["Cantidad_Tareas"] = task_count

    project_df = pd.DataFrame([project_data])
    df = pd.concat([project_df, df], ignore_index=True)

    # Contar tareas por nodo
    for idx, row in df.iterrows():
        depth = wbs_depth(row["WBS"])
        if depth == 3:
            df.at[idx, "Cantidad_Tareas"] = 1
        elif depth is not None:
            wbs_val = normalize_text(row["WBS"])
            child_tasks = df[(df["__depth__"] == 3) & (df["WBS"].map(lambda x: x.startswith(wbs_val + ".") if x else False))]
            df.at[idx, "Cantidad_Tareas"] = len(child_tasks)

    return df, project_name


# ===========================================================================
# Construcción del EDT desde xlsbCFA1.xlsb (reemplaza a xlsb7BAE.xlsb)
# ===========================================================================

def build_wbs_from_fs_cod_part(fs_cod_part):
    """Reconstruye el WBS punteado desde las posiciones 13-20 de FS_COD_PART.
    
    Ejemplo: '090047001PN0202030205' -> componente=02, subcomp=03, act=02, tarea=05 -> '2.3.2.5'
    """
    code = str(fs_cod_part)
    length = len(code)

    if length < 15:
        return None, None

    comp = code[13:15]

    if length == 15:
        return str(int(comp)), "Componente"
    elif length == 17:
        subcomp = code[15:17]
        return f"{int(comp)}.{int(subcomp)}", "Subcomponente"
    elif length == 19:
        subcomp = code[15:17]
        act = code[17:19]
        return f"{int(comp)}.{int(subcomp)}.{int(act)}", "Actividad"
    elif length >= 21:
        subcomp = code[15:17]
        act = code[17:19]
        tarea = code[19:21]
        return f"{int(comp)}.{int(subcomp)}.{int(act)}.{int(tarea)}", "Tarea"
    else:
        return None, None


def build_edt_from_cfa1(df_cfa1, project_code, year=TARGET_YEAR):
    """Construye la estructura EDT desde CFA1 para un proyecto específico.
    
    Filtra filas con FS_COD_PART que comience con el código de proyecto,
    y que tengan longitud >= 15 (Componente en adelante).
    Solo incluye tareas (len=21) con presupuesto planificado > 0.
    """
    df = df_cfa1[df_cfa1["FS_COD_PART"].str.startswith(project_code)].copy()
    df["len"] = df["FS_COD_PART"].str.len()

    # Solo niveles Componente (15) hasta Tarea (21)
    df = df[df["len"].isin([15, 17, 19, 21])].copy()

    if df.empty:
        return pd.DataFrame(columns=["WBS", "Nivel", "Cantidad_Tareas", "Proyecto", "Name"])

    # Calcular presupuesto planificado total por fila para filtrar tareas sin presupuesto
    pres_cols = [c for c in df.columns if c.startswith("FN_IMP_PRES_MEFI_TR")]
    if pres_cols:
        df["__pres_total__"] = df[pres_cols].fillna(0).sum(axis=1)
    else:
        df["__pres_total__"] = 0

    # Excluir tareas (len=21) con presupuesto = 0
    task_zero_mask = (df["len"] == 21) & (df["__pres_total__"] <= 0)
    df = df[~task_zero_mask].copy()

    if df.empty:
        return pd.DataFrame(columns=["WBS", "Nivel", "Cantidad_Tareas", "Proyecto", "Name"])

    # Reconstruir WBS
    wbs_info = df["FS_COD_PART"].apply(build_wbs_from_fs_cod_part)
    df["WBS"] = wbs_info.apply(lambda x: x[0])
    df["Nivel"] = wbs_info.apply(lambda x: x[1])
    df = df[df["WBS"].notna()].copy()

    # Solo conservar ramas que tengan al menos una tarea
    tasks = df[df["len"] == 21].copy()
    if tasks.empty:
        return pd.DataFrame(columns=["WBS", "Nivel", "Cantidad_Tareas", "Proyecto", "Name"])

    valid_wbs = set(tasks["WBS"].tolist())
    for task_wbs in tasks["WBS"].tolist():
        for levels_up in (1, 2, 3):
            p = parent_wbs(task_wbs, levels_up)
            if p:
                valid_wbs.add(p)

    df = df[df["WBS"].isin(valid_wbs)].copy()

    # Deduplicar por WBS (tomar el primer nombre encontrado)
    df = df.drop_duplicates(subset=["WBS"], keep="first")

    # Contar tareas por nodo
    task_count = {wbs: 1 for wbs in tasks["WBS"].tolist()}
    counts_by_wbs = {}
    for wbs, count in task_count.items():
        counts_by_wbs[wbs] = counts_by_wbs.get(wbs, 0) + count
        for levels_up in (1, 2, 3):
            p = parent_wbs(wbs, levels_up)
            if p:
                counts_by_wbs[p] = counts_by_wbs.get(p, 0) + count

    df["Cantidad_Tareas"] = df["WBS"].map(counts_by_wbs).fillna(0).astype(int)
    df["Proyecto"] = project_code
    df["Name"] = df["FS_DES_PART"]

    # Agregar fila raíz de proyecto (Nivel = Proyecto, WBS = "")
    project_row = pd.DataFrame([{
        "WBS": "",
        "Nivel": "Proyecto",
        "Cantidad_Tareas": len(tasks),
        "Proyecto": project_code,
        "Name": project_code
    }])
    
    final_df = pd.concat([project_row, df[["WBS", "Nivel", "Cantidad_Tareas", "Proyecto", "Name"]]], ignore_index=True)
    return final_df.copy()

# ===========================================================================
# Comparación Sciforma vs EDT CFA1 y recálculo con estructura EMMA
# ===========================================================================

def mark_and_recalculate_with_edt(sciforma_df, edt_df):
    """Recalcula avances usando la estructura EDT de CFA1.

    Reglas:
    - Tareas Sciforma que no esten en el EDT de CFA1 quedan con % = NaN.
    - Avance se recalcula por hijos inmediatos validos del EDT:
        Actividad  -> promedio de sus Tareas en EDT.
        Subcomp    -> promedio de sus Actividades en EDT.
        Componente -> promedio de sus Subcomponentes en EDT.
        Proyecto   -> promedio de sus Componentes en EDT.
    - Si un hijo EDT no tiene avance en Sciforma, aporta 0 al promedio.
    """
    df = sciforma_df.copy()
    if df.empty:
        return df

    df["Proyecto"] = df["Proyecto"].map(normalize_text)
    df["WBS"] = df["WBS"].map(normalize_text)
    df["Nivel"] = df["Nivel"].map(normalize_text)
    df["% Comp. Prom."] = pd.to_numeric(df["% Comp. Prom."], errors="coerce")

    if edt_df is None or edt_df.empty:
        df["% Comp. Prom."] = np.nan
        return df

    edt = edt_df.copy()
    edt["Proyecto"] = edt["Proyecto"].map(normalize_text)
    edt["WBS"] = edt["WBS"].map(normalize_text)
    edt["Nivel"] = edt["Nivel"].map(normalize_text)

    # Mapas de EDT
    edt_wbs_by_project = {}
    edt_children = {}       # (project, parent_wbs) -> [child_wbs]
    edt_project_children = {}  # project -> [component_wbs]

    for _, r in edt.iterrows():
        project = normalize_text(r["Proyecto"])
        wbs = normalize_text(r["WBS"])
        nivel = normalize_text(r["Nivel"])
        edt_wbs_by_project.setdefault(project, set()).add(wbs)
        if nivel == "Componente":
            edt_project_children.setdefault(project, []).append(wbs)
        elif nivel in ("Subcomponente", "Actividad", "Tarea"):
            p = parent_wbs(wbs, 1)
            if p:
                edt_children.setdefault((project, p), []).append(wbs)

    # Deduplicar listas preservando orden
    edt_project_children = {k: list(dict.fromkeys(v)) for k, v in edt_project_children.items()}
    edt_children = {k: list(dict.fromkeys(v)) for k, v in edt_children.items()}

    # Marcar presencia
    def _present(row):
        project = normalize_text(row.get("Proyecto", ""))
        nivel = normalize_text(row.get("Nivel", ""))
        wbs = normalize_text(row.get("WBS", ""))
        if nivel == "Proyecto":
            return project in edt_wbs_by_project or extract_project_code(project) in edt_wbs_by_project
        # Buscar por código de proyecto
        proj_code = extract_project_code(project)
        return wbs in edt_wbs_by_project.get(proj_code, set())

    df["__present__"] = df.apply(_present, axis=1)

    # Anular avance de tareas no presentes en EDT
    task_mask = df["Nivel"].eq("Tarea")
    df.loc[task_mask & ~df["__present__"], "% Comp. Prom."] = np.nan

    # Obtener progreso de tareas Sciforma que están en EDT
    task_progress = {}
    valid_tasks = df[task_mask & df["__present__"]].copy()
    for _, task in valid_tasks.iterrows():
        project = extract_project_code(normalize_text(task.get("Proyecto", "")))
        wbs = normalize_text(task.get("WBS", ""))
        progress = task.get("% Comp. Prom.")
        task_progress[(project, wbs)] = float(progress) if pd.notna(progress) else 0.0

    # Recalcular avances en cascada según estructura EDT
    value_by_node = {}
    value_by_project = {}

    for project in sorted(edt_wbs_by_project):
        project_wbs_set = edt_wbs_by_project.get(project, set())

        # Tareas
        for wbs in [w for w in project_wbs_set if wbs_depth(w) == 3]:
            value_by_node[(project, wbs)] = task_progress.get((project, wbs), 0.0)

        # Actividades (promedio de sus tareas EDT)
        for wbs in sorted([w for w in project_wbs_set if wbs_depth(w) == 2]):
            children = edt_children.get((project, wbs), [])
            if children:
                value_by_node[(project, wbs)] = sum(
                    value_by_node.get((project, c), 0.0) for c in children
                ) / len(children)

        # Subcomponentes (promedio de sus actividades EDT)
        for wbs in sorted([w for w in project_wbs_set if wbs_depth(w) == 1]):
            children = edt_children.get((project, wbs), [])
            if children:
                value_by_node[(project, wbs)] = sum(
                    value_by_node.get((project, c), 0.0) for c in children
                ) / len(children)

        # Componentes (promedio de sus subcomponentes EDT)
        for wbs in sorted([w for w in project_wbs_set if wbs_depth(w) == 0]):
            children = edt_children.get((project, wbs), [])
            if children:
                value_by_node[(project, wbs)] = sum(
                    value_by_node.get((project, c), 0.0) for c in children
                ) / len(children)

        # Proyecto (promedio de componentes)
        proj_children = edt_project_children.get(project, [])
        if proj_children:
            value_by_project[project] = sum(
                value_by_node.get((project, c), 0.0) for c in proj_children
            ) / len(proj_children)

    # Aplicar valores recalculados a la estructura completa (edt)
    edt["% Comp. Prom."] = np.nan
    for idx, row in edt.iterrows():
        project = extract_project_code(normalize_text(row.get("Proyecto", "")))
        wbs = normalize_text(row.get("WBS", ""))
        nivel = normalize_text(row.get("Nivel", ""))
        
        if nivel.lower() == "proyecto":
            if project in value_by_project:
                edt.at[idx, "% Comp. Prom."] = value_by_project[project]
        else:
            if (project, wbs) in value_by_node:
                edt.at[idx, "% Comp. Prom."] = value_by_node[(project, wbs)]
                
    return edt

# ===========================================================================
# Función principal para construir Fact_Avance_Fisico
# ===========================================================================

def build_fact_avance_fisico(sciforma_dir, df_cfa1, dim_tecnica, year=TARGET_YEAR):
    """Construye la tabla de hechos Fact_Avance_Fisico.
    
    1. Lee todos los archivos Sciforma
    2. Lee la estructura EDT desde CFA1 por proyecto
    3. Compara y recalcula avances
    4. Mapea a ID_Tecnica exacto por nivel (Componente, Subcomponente, etc.)
    """
    sciforma_path = Path(sciforma_dir)
    sciforma_files = sorted(sciforma_path.glob("sciforma-print*.xlsx"))
    sciforma_files = [f for f in sciforma_files if not f.name.startswith("~$")]

    if not sciforma_files:
        print(f"  ADVERTENCIA: No se encontraron archivos sciforma-print*.xlsx en {sciforma_dir}")
        return pd.DataFrame()

    print(f"  Encontrados {len(sciforma_files)} archivos Sciforma.")

    # 1. Procesar cada archivo Sciforma
    all_sciforma = []
    for file_path in sciforma_files:
        try:
            df, project_name = process_one_sciforma_file(file_path, year)
            all_sciforma.append(df)
            print(f"    [OK] {file_path.name} -> {project_name}")
        except Exception as exc:
            print(f"    [X] {file_path.name} -> ERROR: {exc}")

    if not all_sciforma:
        print("  ADVERTENCIA: No se pudo procesar ningún archivo Sciforma.")
        return pd.DataFrame()

    consolidated = pd.concat(all_sciforma, ignore_index=True)

    # 2. Obtener códigos de proyecto únicos desde Sciforma
    project_names = consolidated[consolidated["Nivel"] == "Proyecto"]["Proyecto"].unique()
    project_codes = {extract_project_code(name): name for name in project_names if extract_project_code(name)}

    print(f"  Proyectos Sciforma detectados: {list(project_codes.keys())}")

    # 3. Construir EDT desde CFA1 para cada proyecto
    all_edt = []
    for code in sorted(project_codes.keys()):
        edt = build_edt_from_cfa1(df_cfa1, code, year)
        if not edt.empty:
            all_edt.append(edt)
            task_count = len(edt[edt["Nivel"] == "Tarea"])
            print(f"    EDT {code}: {len(edt)} nodos, {task_count} tareas con presupuesto")

    edt_consolidated = pd.concat(all_edt, ignore_index=True) if all_edt else pd.DataFrame()

    # 4. Comparar y recalcular avances
    consolidated = mark_and_recalculate_with_edt(consolidated, edt_consolidated)

    # 5. Mapear a ID_Tecnica
    # dim_tecnica tiene Proyecto y WBS, así que la búsqueda es 1:1 para TODOS los niveles
    tecnica_to_id = dict(zip(zip(dim_tecnica['Proyecto'], dim_tecnica['WBS']), dim_tecnica['ID_Tecnica']))

    # Construir Fact_Avance_Fisico
    fact_rows = []
    
    # ID de trimestres basados en TARGET_YEAR
    trimestres = {
        "TR01": year * 10 + 1,
        "TR02": year * 10 + 2,
        "TR03": year * 10 + 3,
        "TR04": year * 10 + 4
    }
    
    for _, row in consolidated.iterrows():
        project = normalize_text(row.get("Proyecto", ""))
        proj_code = extract_project_code(project)
        nivel = normalize_text(row.get("Nivel", ""))
        wbs = normalize_text(row.get("WBS", ""))
        pct_avance = row.get("% Comp. Prom.")

        # ── Modelo Kimball Puro: solo exportar el nivel más granular (Tarea) ──
        # Los niveles superiores (Actividad, Subcomponente, Componente, Proyecto)
        # serán calculados por DAX en Power BI usando AVERAGEX anidados.
        if nivel.lower() != "tarea":
            continue

        if pd.isna(pct_avance):
            pct_avance = 0.0  # Tareas sin avance en Sciforma = 0% (EMMA es la fuente de verdad)

        # Buscar ID_Tecnica exacto
        id_tecnica = tecnica_to_id.get((proj_code, wbs))

        if id_tecnica:
            # Crear 4 filas por nodo (una por trimestre)
            for tr_str, id_trim in trimestres.items():
                fact_rows.append({
                    "ID_Tecnica": id_tecnica,
                    "ID_Trimestre": id_trim,
                    "Meta_Fisica": 0.25, # Meta repartida equitativamente
                    # Asignamos el avance físico real solo al trimestre en curso
                    "Pct_Avance_Fisico": float(pct_avance) if (tr_str == CURRENT_QUARTER) else None,
                })

    fact_df = pd.DataFrame(fact_rows)

    return fact_df


# ===========================================================================
# ETL Principal
# ===========================================================================

def run_etl():
    print("=" * 60)
    print(f"INICIANDO ETL - MODELO ESTRELLA BI {TARGET_YEAR}")
    print("=" * 60)

    script_dir = os.path.dirname(os.path.abspath(__file__))
    data_dir = os.path.join(os.path.dirname(script_dir), "data")
    
    xlsb_files = glob.glob(os.path.join(data_dir, "*.xlsb"))
    if not xlsb_files:
        raise FileNotFoundError(f"No se encontró ningún archivo .xlsb en la carpeta {data_dir}")
    
    file_path = xlsb_files[0]
    print(f"  [!] Archivo ERP detectado: {os.path.basename(file_path)}")

    # 1. Cargar Datos
    print("\n[1/6] Cargando archivo plano...")
    df = pd.read_excel(file_path, engine='pyxlsb', sheet_name='tmpCBA8')

    # Filtrar posibles nulos en la clave (aunque según el usuario viene consistente del ERP)
    df = df[df['FS_COD_PART'].notna()]
    df['FS_COD_PART'] = df['FS_COD_PART'].astype(str)

    # 2. Dim_Administrativa y Dim_Estructura_WBS
    print("\\n[2/6] Generando Dimensiones (Admin y WBS)...")
    
    def safe_slice(val, start, end):
        return val[start:end] if len(val) >= start else ""

    def prefix_if_long_enough(val, length):
        return val[0:length] if len(val) >= length else None

    desc_dict = dict(zip(df['FS_COD_PART'], df['FS_DES_PART']))

    # --- Dim_Administrativa ---
    df['Admin_Code'] = df['FS_COD_PART'].apply(lambda x: x[:13] if len(x) >= 13 else None)
    admin_codes = df['Admin_Code'].dropna().drop_duplicates().reset_index(drop=True)
    dim_admin = pd.DataFrame({'Codigo_Admin': admin_codes})
    
    dim_admin['Proyecto'] = dim_admin['Codigo_Admin'].apply(lambda x: safe_slice(x, 0, 3))
    dim_admin['Proyecto_Desc'] = dim_admin['Codigo_Admin'].apply(lambda x: prefix_if_long_enough(x, 3)).map(desc_dict)
    dim_admin['Donante'] = dim_admin['Codigo_Admin'].apply(lambda x: safe_slice(x, 3, 6))
    dim_admin['Donante_Desc'] = dim_admin['Codigo_Admin'].apply(lambda x: prefix_if_long_enough(x, 6)).map(desc_dict)
    dim_admin['Ejecutor'] = dim_admin['Codigo_Admin'].apply(lambda x: safe_slice(x, 6, 9))
    dim_admin['Ejecutor_Desc'] = dim_admin['Codigo_Admin'].apply(lambda x: prefix_if_long_enough(x, 9)).map(desc_dict)
    dim_admin['Unidad_Ejecutora'] = dim_admin['Codigo_Admin'].apply(lambda x: safe_slice(x, 9, 13))
    dim_admin['Unidad_Ejecutora_Desc'] = dim_admin['Codigo_Admin'].apply(lambda x: prefix_if_long_enough(x, 13)).map(desc_dict)
    
    dim_admin.insert(0, 'ID_Admin', range(1, len(dim_admin) + 1))
    map_admin = dict(zip(dim_admin['Codigo_Admin'], dim_admin['ID_Admin']))
    print(f"  Dim_Administrativa: {len(dim_admin)} registros")

    # --- Dim_Tecnica ---
    def build_wbs_str(code):
        length = len(code)
        if length < 15: return ""
        parts = []
        for i in range(13, min(length, 21), 2):
            parts.append(str(int(code[i:i+2])))
        return ".".join(parts)

    def get_wbs_level(code):
        length = len(code)
        if length == 3: return 'Proyecto'
        if length == 15: return 'Componente'
        if length == 17: return 'Subcomponente'
        if length == 19: return 'Actividad'
        if length >= 21: return 'Tarea'
        return 'Desconocido'

    df['Phys_Code'] = df['FS_COD_PART'].apply(lambda x: x[:3] + '-' + x[13:21] if len(x) >= 13 else x[:3])
    df['WBS'] = df['FS_COD_PART'].apply(build_wbs_str)

    phys_codes = df[['Phys_Code', 'FS_COD_PART', 'WBS', 'FS_DES_PART']].copy()
    phys_codes['Length'] = phys_codes['FS_COD_PART'].str.len()
    phys_codes = phys_codes.sort_values('Length')
    phys_codes = phys_codes.drop_duplicates(subset=['Phys_Code'], keep='first').reset_index(drop=True)

    dim_tecnica = pd.DataFrame({
        'Codigo_Tecnico': phys_codes['Phys_Code'],
        'Original_Code': phys_codes['FS_COD_PART'],
        'WBS': phys_codes['WBS'],
        'Nivel': phys_codes['FS_COD_PART'].apply(get_wbs_level),
        'Descripcion': phys_codes['FS_DES_PART']
    })
    dim_tecnica = dim_tecnica[dim_tecnica['Nivel'] != 'Desconocido'].copy()
    
    # Desglose de nombres
    dim_tecnica['Proyecto'] = dim_tecnica['Original_Code'].apply(lambda x: safe_slice(x, 0, 3))
    dim_tecnica['Proyecto_Desc'] = dim_tecnica['Original_Code'].apply(lambda x: prefix_if_long_enough(x, 3)).map(desc_dict)
    dim_tecnica['Componente'] = dim_tecnica['Original_Code'].apply(lambda x: safe_slice(x, 13, 15))
    dim_tecnica['Componente_Desc'] = dim_tecnica['Original_Code'].apply(lambda x: prefix_if_long_enough(x, 15)).map(desc_dict)
    dim_tecnica['Subcomponente'] = dim_tecnica['Original_Code'].apply(lambda x: safe_slice(x, 15, 17))
    dim_tecnica['Subcomponente_Desc'] = dim_tecnica['Original_Code'].apply(lambda x: prefix_if_long_enough(x, 17)).map(desc_dict)
    dim_tecnica['Actividad'] = dim_tecnica['Original_Code'].apply(lambda x: safe_slice(x, 17, 19))
    dim_tecnica['Actividad_Desc'] = dim_tecnica['Original_Code'].apply(lambda x: prefix_if_long_enough(x, 19)).map(desc_dict)
    dim_tecnica['Tarea'] = dim_tecnica['Original_Code'].apply(lambda x: safe_slice(x, 19, 21))
    dim_tecnica['Tarea_Desc'] = dim_tecnica['Original_Code'].apply(lambda x: prefix_if_long_enough(x, 21)).map(desc_dict)
    
    dim_tecnica.drop(columns=['Original_Code'], inplace=True)
    dim_tecnica.insert(0, 'ID_Tecnica', range(1, len(dim_tecnica) + 1))
    print(f"  Dim_Tecnica: {len(dim_tecnica)} registros")

    # 3. Dim_Tiempo
    print("\n[3/6] Generando Dim_Tiempo...")
    dim_tiempo = pd.DataFrame({
        'ID_Trimestre': [TARGET_YEAR * 10 + 1, TARGET_YEAR * 10 + 2, TARGET_YEAR * 10 + 3, TARGET_YEAR * 10 + 4],
        'Trimestre_Texto': ['TR01', 'TR02', 'TR03', 'TR04'],
        'Año': [TARGET_YEAR, TARGET_YEAR, TARGET_YEAR, TARGET_YEAR],
        'Trimestre_Numero': [1, 2, 3, 4],
        'Fecha_Fin_Trimestre': [f'{TARGET_YEAR}-03-31', f'{TARGET_YEAR}-06-30', f'{TARGET_YEAR}-09-30', f'{TARGET_YEAR}-12-31']
    })
    map_tiempo = {'TR01': TARGET_YEAR * 10 + 1, 'TR02': TARGET_YEAR * 10 + 2, 'TR03': TARGET_YEAR * 10 + 3, 'TR04': TARGET_YEAR * 10 + 4}

    # 4. Fact_Ejecucion (Unpivot/Melt)
    print("\n[4/6] Generando Fact_Ejecucion...")
    df_facts = df[df['FS_COD_PART'].str.len() == 21].copy()

    # Monto total original para verificación
    presupuesto_total_original = df_facts[['FN_IMP_PRES_MEFI_TR01', 'FN_IMP_PRES_MEFI_TR02', 'FN_IMP_PRES_MEFI_TR03', 'FN_IMP_PRES_MEFI_TR04']].sum().sum()

    def melt_metric(df_in, prefix, value_name):
        cols = [f"{prefix}_TR01", f"{prefix}_TR02", f"{prefix}_TR03", f"{prefix}_TR04"]
        m = df_in[['FS_COD_PART'] + cols].melt(id_vars='FS_COD_PART', value_vars=cols, var_name='Trimestre', value_name=value_name)
        m['Trimestre_Texto'] = m['Trimestre'].str[-4:]
        m.drop(columns=['Trimestre'], inplace=True)
        return m

    df_presupuesto = melt_metric(df_facts, 'FN_IMP_PRES_MEFI', 'Presupuesto_Financiero')
    df_ejecucion_fin = melt_metric(df_facts, 'FN_IMP_EJEC_MEFI', 'Ejecucion_Financiera')

    fact = df_presupuesto.merge(df_ejecucion_fin, on=['FS_COD_PART', 'Trimestre_Texto'])

    fact['Admin_Code'] = fact['FS_COD_PART'].apply(lambda x: x[:13])
    fact['Phys_Code'] = fact['FS_COD_PART'].apply(lambda x: x[:3] + '-' + x[13:21])
    
    fact['ID_Admin'] = fact['Admin_Code'].map(map_admin)
    # Map ID_Tecnica using Codigo_Tecnico
    map_tecnica = dict(zip(dim_tecnica['Codigo_Tecnico'], dim_tecnica['ID_Tecnica']))
    fact['ID_Tecnica'] = fact['Phys_Code'].map(map_tecnica)
    fact['ID_Trimestre'] = fact['Trimestre_Texto'].map(map_tiempo)

    fact = fact[['ID_Admin', 'ID_Tecnica', 'ID_Trimestre', 'Presupuesto_Financiero', 'Ejecucion_Financiera']]

    # Validacion
    presupuesto_total_fact = fact['Presupuesto_Financiero'].sum()
    print(f"  Validación Presupuesto Financiero: Origen = {presupuesto_total_original:.2f} | Fact = {presupuesto_total_fact:.2f}")
    if abs(presupuesto_total_original - presupuesto_total_fact) > 0.01:
        print("  [!] ADVERTENCIA: Discrepancia en los montos de Presupuesto!")
    else:
        print("  [OK] VALIDACIÓN EXITOSA: Los montos coinciden perfectamente.")

    print(f"  Fact_Ejecucion: {len(fact)} registros")

    # 5. Fact_Avance_Fisico
    print("\n[5/6] Generando Fact_Avance_Fisico...")
    sciforma_dir = os.path.join(os.path.dirname(script_dir), "data", "Fisico")

    if os.path.exists(sciforma_dir):
        fact_avance = build_fact_avance_fisico(sciforma_dir, df, dim_tecnica, year=TARGET_YEAR)
        print(f"  Fact_Avance_Fisico: {len(fact_avance)} registros")

        # Validación: verificar cantidad de tareas por proyecto
        if not fact_avance.empty:
            fact_merged = fact_avance.merge(dim_tecnica[['ID_Tecnica', 'Nivel', 'Proyecto']], on='ID_Tecnica')
            for proj_code in sorted(fact_merged['Proyecto'].unique()):
                task_count = len(fact_merged[(fact_merged['Proyecto'] == proj_code) & (fact_merged['Nivel'] == 'Tarea')]['ID_Tecnica'].unique())
                print(f"    Proyecto {proj_code}: {task_count} tareas en Fact_Avance_Fisico")
    else:
        print(f"  [!] Directorio '{sciforma_dir}' no encontrado. Se omite Fact_Avance_Fisico.")
        fact_avance = pd.DataFrame()

    # 6. Exportar
    print("\n[6/6] Exportando modelo_estrella.xlsx y data_poc.json...")
    
    # Asegurar que exista la carpeta output/ en la raíz del proyecto
    output_dir = os.path.join(os.path.dirname(script_dir), "output")
    os.makedirs(output_dir, exist_ok=True)
    
    out_excel = os.path.join(output_dir, "modelo_estrella.xlsx")
    with pd.ExcelWriter(out_excel, engine='openpyxl') as writer:
        fact.to_excel(writer, sheet_name='Fact_Ejecucion', index=False)
        dim_admin.to_excel(writer, sheet_name='Dim_Administrativa', index=False)
        
        # ── Modelo Kimball: exportar solo filas de nivel Tarea y eliminar columna Nivel ──
        dim_tecnica_export = dim_tecnica[dim_tecnica['Nivel'] == 'Tarea'].copy()
        dim_tecnica_export = dim_tecnica_export.drop(columns=['WBS', 'Descripcion', 'Nivel'], errors='ignore')
        dim_tecnica_export.to_excel(writer, sheet_name='Dim_Tecnica', index=False)
        dim_tiempo.to_excel(writer, sheet_name='Dim_Tiempo', index=False)
        if not fact_avance.empty:
            fact_avance.to_excel(writer, sheet_name='Fact_Avance_Fisico', index=False)

    # JSON para React PoC
    dashboard_data = fact.merge(dim_tiempo[['ID_Trimestre', 'Trimestre_Texto']], on='ID_Trimestre')
    resumen_trimestral = dashboard_data.groupby('Trimestre_Texto').agg({
        'Presupuesto_Financiero': 'sum',
        'Ejecucion_Financiera': 'sum'
    }).reset_index()

    kpi_global = {
        'Presupuesto_Total': float(dashboard_data['Presupuesto_Financiero'].sum()),
        'Ejecucion_Financiera_Total': float(dashboard_data['Ejecucion_Financiera'].sum())
    }

    # Resumen financiero y físico por Proyecto
    resumen_proyecto_json = []
    if not fact_avance.empty:
        current_trim_id = TARGET_YEAR * 10 + int(CURRENT_QUARTER[-1])
        
        # 1. Avance Físico del trimestre actual
        fact_actual = fact_avance[fact_avance['ID_Trimestre'] == current_trim_id]
        fact_merged = fact_actual.merge(dim_tecnica[['ID_Tecnica', 'Nivel', 'Proyecto']], on='ID_Tecnica')
        proyectos_fisico = fact_merged[fact_merged["Nivel"] == "Proyecto"].copy()
        
        # 2. Avance Financiero acumulado por proyecto
        fin_merged = fact.merge(dim_tecnica[['ID_Tecnica', 'Proyecto']], on='ID_Tecnica')
        fin_agrupado = fin_merged.groupby('Proyecto').agg({
            'Presupuesto_Financiero': 'sum',
            'Ejecucion_Financiera': 'sum'
        }).reset_index()
        
        # 3. Combinar ambos
        for _, row in proyectos_fisico.iterrows():
            proyecto = row.get("Proyecto", "")
            fisico_val = float(row["Pct_Avance_Fisico"]) if pd.notna(row.get("Pct_Avance_Fisico")) else 0
            
            # Buscar info financiera
            fin_info = fin_agrupado[fin_agrupado['Proyecto'] == proyecto]
            presupuesto = float(fin_info['Presupuesto_Financiero'].iloc[0]) if not fin_info.empty else 0.0
            ejecucion = float(fin_info['Ejecucion_Financiera'].iloc[0]) if not fin_info.empty else 0.0
            
            pct_financiero = (ejecucion / presupuesto) if presupuesto > 0 else 0.0
            
            resumen_proyecto_json.append({
                "Proyecto": proyecto,
                "Pct_Avance_Fisico": fisico_val,
                "Presupuesto_Financiero": presupuesto,
                "Ejecucion_Financiera": ejecucion,
                "Pct_Avance_Financiero": pct_financiero,
                "Brecha_Fisica": 1.0 - fisico_val,
                "Brecha_Financiera": 1.0 - pct_financiero
            })

    json_out = {
        'kpis': kpi_global,
        'trimestres': resumen_trimestral.to_dict(orient='records'),
        'resumen_proyectos': resumen_proyecto_json
    }

    out_json = os.path.join(output_dir, "data_poc.json")
    with open(out_json, "w", encoding='utf-8') as f:
        json.dump(json_out, f, indent=4, ensure_ascii=False)
        
    print(f"  [OK] Archivos guardados exitosamente en la carpeta: {output_dir}")

    print("\n" + "=" * 60)
    print("ETL COMPLETADO EXITOSAMENTE")
    print("=" * 60)
    sheets = ["Fact_Ejecucion", "Dim_Administrativa", "Dim_Tecnica", "Dim_Tiempo"]
    if not fact_avance.empty:
        sheets.append("Fact_Avance_Fisico")
    print(f"  -> modelo_estrella.xlsx ({', '.join(sheets)})")
    print(f"  -> data_poc.json")


if __name__ == "__main__":
    run_etl()
