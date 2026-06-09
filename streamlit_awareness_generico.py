# -*- coding: utf-8 -*-
"""
Generador Flexible de Awareness / Recordación - Streamlit

Funciona para:
- Bancos
- Conglomerados Financieros (Nueva Base)
- Marcas / Empresas / Productos / Personalizado
"""

import io
import re
import json
import unicodedata
from typing import Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ============================================================
# CONFIGURACIONES PREDEFINIDAS
# ============================================================

# --- PRESET 1: BANCOS (Original) ---
BANKS = [
    "Banco Agrario",
    "Bancolombia",
    "Davivienda",
    "Banco de Bogotá",
    "BBVA",
    "Banco Caja Social",
    "Banco Popular",
    "Bancamía",
    "Banco de Occidente",
    "Banco Mundo Mujer",
]

BANK_ALIASES = {
    "Banco Agrario": ["banco agrario", "banco agrario de colombia", "agrario"],
    "Bancolombia": ["bancolombia"],
    "Davivienda": ["davivienda"],
    "Banco de Bogotá": ["banco de bogota", "banco de bogotá", "banco bogota", "banco bogotá"],
    "BBVA": ["bbva", "bbvva", "bvva", "bbwa", "bva", "bvvwa"],
    "Banco Caja Social": ["banco caja social", "caja social", "cajas social", "caja sosial"],
    "Banco Popular": ["banco popular", "popular"],
    "Bancamía": ["bancamia", "bancamía", "banca mia", "banca mía"],
    "Banco de Occidente": ["banco de occidente", "occidente"],
    "Banco Mundo Mujer": [
        "banco mundo mujer",
        "mundo mujer",
        "banco de la mujer",
        "fundacion de la mujer",
        "fundación de la mujer",
        "de la mujer",
    ],
}

BANK_NORMALIZATIONS = [
    {"pattern": r"(?i)Banco\s+de\s+Bogota", "replacement": "Banco de Bogotá"},
    {"pattern": r"(?i)(?<!Banco\s)Caja\s+Social", "replacement": "Banco Caja Social"},
    {"pattern": r"(?i)Bancamia", "replacement": "Bancamía"},
]

# --- PRESET 2: CONGLOMERADOS FINANCIEROS (Nueva Base) ---
CONGLOMERATES = [
    "Conglomerado BBVA",
    "Grupo cooperativo Coomeva",
    "Fundación Grupo Social",
    "Grupo Bolivar",
    "Conglomerado financiero Sura-Bancolombia",
    "Grupo Aval",
    "GNB Sudameris",
    "Conglomerado Credicorp capital"
]

CONGLOMERATE_ALIASES = {
    "Conglomerado BBVA": ["conglomerado bbva", "bbva"],
    "Grupo cooperativo Coomeva": ["grupo cooperativo coomeva", "coomeva"],
    "Fundación Grupo Social": ["fundacion grupo social", "grupo social"],
    "Grupo Bolivar": ["grupo bolivar", "bolivar"],
    "Conglomerado financiero Sura-Bancolombia": ["conglomerado financiero sura bancolombia", "sura", "bancolombia"],
    "Grupo Aval": ["grupo aval", "aval"],
    "GNB Sudameris": ["gnb sudameris", "sudameris", "gnb"],
    "Conglomerado Credicorp capital": ["conglomerado credicorp capital", "credicorp"]
}

NEGATIVOS = {
    "", "0", "no", "nan", "none", "false", "ninguno", "ninguna",
    "no se", "no sé", "no recuerdo", "ningun otro", "ningún otro", "no aplica",
}


# ============================================================
# UTILIDADES DE TEXTO
# ============================================================

def norm(value) -> str:
    """Normaliza texto para comparar sin tildes, símbolos ni mayúsculas."""
    if pd.isna(value):
        return ""
    text = str(value).strip().lower()
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
    text = re.sub(r"[^a-z0-9\s]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()

def safe_sheet_name(name: str) -> str:
    """Excel limita nombres de hoja a 31 caracteres."""
    cleaned = re.sub(r"[\\/*?:\[\]]", " ", str(name)).strip()
    return cleaned[:31] or "Hoja"

def apply_normalizations(value, normalizations: List[Dict]) -> str:
    """Aplica normalizaciones editables por el usuario."""
    if pd.isna(value):
        return value
    text = str(value)
    for item in normalizations or []:
        pattern = item.get("pattern", "")
        replacement = item.get("replacement", "")
        if not pattern:
            continue
        try:
            text = re.sub(pattern, replacement, text)
        except re.error:
            pass
    return text

def contains_entity(value, entity: str, aliases: Dict[str, List[str]]) -> bool:
    """Valida si una respuesta contiene una entidad o alguno de sus alias."""
    text = norm(value)
    if text in NEGATIVOS:
        return False
    for alias in aliases.get(entity, [entity]):
        alias_norm = norm(alias)
        if alias_norm and re.search(r"(^|\s)" + re.escape(alias_norm) + r"(\s|$)", text):
            return True
    return False


# ============================================================
# DETECCIÓN DE COLUMNAS (MEJORADA PARA MATRICES)
# ============================================================

def find_col(df: pd.DataFrame, prefix: str) -> Optional[str]:
    """Encuentra la primera columna cuyo encabezado empieza por el prefijo."""
    prefix_norm = norm(prefix)
    if not prefix_norm or prefix_norm == "0":
        return None
    for column in df.columns:
        if norm(column).startswith(prefix_norm):
            return column
    return None

def prefixes_to_list(text: str) -> List[str]:
    """Convierte un textarea en lista de prefijos, uno por línea."""
    return [line.strip() for line in str(text).splitlines() if line.strip()]

def find_cols(df: pd.DataFrame, prefixes_text: str) -> List[str]:
    """
    Encuentra TODAS las columnas que empiecen con alguno de los prefijos.
    Ideal para agrupar todas las columnas de una pregunta tipo matriz (Ej: P2-P2.)
    """
    detected = []
    prefixes = [norm(p) for p in prefixes_to_list(prefixes_text) if p.strip() != "0"]
    
    if not prefixes:
        return detected
        
    for column in df.columns:
        col_norm = norm(column)
        # Si la columna empieza con CUALQUIERA de los prefijos en la lista
        if any(col_norm.startswith(prefix) for prefix in prefixes):
            if column not in detected:
                detected.append(column)
                
    return detected

def entity_from_aided_col(column_name: str, entities: List[str]) -> Optional[str]:
    """Detecta la entidad según el texto del encabezado."""
    text = norm(str(column_name).replace("\n", " "))
    for entity in entities:
        if norm(entity) in text:
            return entity
    return None


# ============================================================
# CÁLCULO DE AWARENESS
# ============================================================

def build_analysis(
    df: pd.DataFrame,
    demo_cols: Dict[str, Optional[str]],
    cfg: Dict,
    entities: List[str],
    aliases: Dict[str, List[str]],
    normalizations: List[Dict],
    expected_raw_cols: int,
) -> Tuple[pd.DataFrame, pd.DataFrame, List[str]]:
    """Construye la hoja base y los indicadores TOM/Espontáneo/Ayudado."""
    tom_col = find_col(df, cfg["tom"])
    esp_cols = find_cols(df, cfg["esp"])
    ayud_cols = find_cols(df, cfg["ayu"])

    raw_cols = [
        demo_cols.get("sexo"),
        demo_cols.get("edad"),
        demo_cols.get("departamento"),
        demo_cols.get("estrato"),
        demo_cols.get("ingreso"),
        tom_col,
    ] + esp_cols + ayud_cols

    raw_cols = [col for col in raw_cols if col is not None]

    if expected_raw_cols and int(expected_raw_cols) > 0 and len(raw_cols) != int(expected_raw_cols):
        raise ValueError(
            f"{cfg['name']}: se esperaban {expected_raw_cols} columnas crudas y se detectaron {len(raw_cols)}. "
            f"Por favor coloca el número 0 en 'Columnas crudas esperadas' en la barra lateral para saltar este error.\n"
            f"Columnas detectadas: {raw_cols}"
        )

    if not tom_col:
        raise ValueError(f"{cfg['name']}: no se detectó la columna TOM.")
    if not esp_cols:
        raise ValueError(f"{cfg['name']}: no se detectaron columnas espontáneas.")
    if not ayud_cols:
        raise ValueError(f"{cfg['name']}: no se detectaron columnas ayudadas.")

    raw_df = df[raw_cols].apply(lambda col: col.map(lambda value: apply_normalizations(value, normalizations)))
    output_df = raw_df.copy()
    indicators = {}

    aided_map = {entity: None for entity in entities}
    for col in ayud_cols:
        entity = entity_from_aided_col(col, entities)
        if entity in aided_map:
            aided_map[entity] = col

    for entity in entities:
        tom = raw_df[tom_col].apply(lambda value: contains_entity(value, entity, aliases))

        espontaneo = pd.Series(False, index=df.index)
        for col in esp_cols:
            espontaneo = espontaneo | raw_df[col].apply(lambda value: contains_entity(value, entity, aliases))

        aided_col = aided_map.get(entity)
        if aided_col is not None:
            def check_aided_value(value, ent, als):
                if pd.isna(value):
                    return False
                v_str = str(value).strip().lower()
                if v_str in ["1", "si", "sí", "x", "seleccionado", "seleccionada", "true"] or ent.lower() in v_str:
                    return True
                return contains_entity(value, ent, als)
            
            ayudado = raw_df[aided_col].apply(lambda value: check_aided_value(value, entity, aliases))
        else:
            ayudado = pd.Series(False, index=df.index)
            for col in ayud_cols:
                ayudado = ayudado | raw_df[col].apply(lambda value: contains_entity(value, entity, aliases))

        indicators[(entity, "TOM")] = tom.astype(int)
        indicators[(entity, "Espontaneo")] = ((~tom) & espontaneo).astype(int)
        indicators[(entity, "Ayudado")] = ((~tom) & (~espontaneo) & ayudado).astype(int)

        output_df[(entity, "TOM")] = indicators[(entity, "TOM")]
        output_df[(entity, "Espontaneo")] = indicators[(entity, "Espontaneo")]
        output_df[(entity, "Ayudado")] = indicators[(entity, "Ayudado")]

    return output_df, pd.DataFrame(indicators), raw_cols


# ============================================================
# TABLAS RESUMEN Y ESCRITURA
# ============================================================

def category_table(indicators: pd.DataFrame, df: pd.DataFrame, category_col: Optional[str], category_value, entities: List[str]) -> pd.DataFrame:
    if category_col is None:
        mask = pd.Series(True, index=df.index)
    else:
        mask = df[category_col].fillna("Sin dato") == category_value

    base = int(mask.sum())
    rows = []

    for entity in entities:
        tom = int(indicators[(entity, "TOM")][mask].sum()) if (entity, "TOM") in indicators else 0
        esp = int(indicators[(entity, "Espontaneo")][mask].sum()) if (entity, "Espontaneo") in indicators else 0
        ayu = int(indicators[(entity, "Ayudado")][mask].sum()) if (entity, "Ayudado") in indicators else 0
        awa = tom + esp + ayu

        rows.append({
            "Entidad": entity,
            "TOM": tom,
            "TOM %": tom / base if base else 0,
            "Espontaneo": esp,
            "Espontaneo %": esp / base if base else 0,
            "Ayudado": ayu,
            "Ayudado %": ayu / base if base else 0,
            "AWA": awa,
            "AWA %": awa / base if base else 0,
        })
    return pd.DataFrame(rows)


def awareness_sections(indicators: pd.DataFrame, df: pd.DataFrame, demo_cols: Dict, analysis_name: str, entities: List[str]):
    sections = [(f"Resumen general {analysis_name}", category_table(indicators, df, None, None, entities))]
    for key in ["sexo", "estrato"]:
        col = demo_cols.get(key)
        if col:
            for value in pd.Series(df[col].fillna("Sin dato")).drop_duplicates():
                sections.append((f"Detalle por {key} - {value}", category_table(indicators, df, col, value, entities)))
    return sections


def department_sections(indicators: pd.DataFrame, df: pd.DataFrame, demo_cols: Dict, entities: List[str]):
    dept_col = demo_cols.get("departamento")
    if not dept_col:
        return [("Sin departamento/ciudad detectado", pd.DataFrame())]
    return [(str(value), category_table(indicators, df, dept_col, value, entities)) 
            for value in pd.Series(df[dept_col].fillna("Sin dato")).drop_duplicates()]


def demographic_sections(df: pd.DataFrame, demo_cols: Dict):
    sections = [("Base", pd.DataFrame({"Indicador": ["Total encuestas procesadas"], "Valor": [len(df)]}))]
    for key, label in [("sexo", "Sexo"), ("edad", "Edad"), ("departamento", "Ciudad/Depto"), ("estrato", "Estrato"), ("ingreso", "Ingreso")]:
        col = demo_cols.get(key)
        if col:
            table = df[col].fillna("Sin dato").value_counts().rename_axis(label).reset_index(name="Cantidad")
            table["%"] = table["Cantidad"] / len(df) if len(df) else 0
            sections.append((f"Demográfico - {label}", table))
    return sections


def style_sheet(ws):
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for row in ws.iter_rows():
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for col_idx in range(1, ws.max_column + 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = 18


def write_main_sheet(wb: Workbook, sheet_name: str, output_df: pd.DataFrame):
    ws = wb.create_sheet(safe_sheet_name(sheet_name))
    for col_idx, column in enumerate(output_df.columns, start=1):
        ws.cell(1, col_idx).value = str(column)
        ws.cell(1, col_idx).fill = PatternFill("solid", fgColor="D9EAF7")
        ws.cell(1, col_idx).font = Font(bold=True)
    for row_idx, row in enumerate(output_df.itertuples(index=False), start=2):
        for col_idx, value in enumerate(row, start=1):
            ws.cell(row_idx, col_idx).value = value
    style_sheet(ws)


def write_sections_sheet(wb: Workbook, sheet_name: str, sections):
    ws = wb.create_sheet(safe_sheet_name(sheet_name))
    row_idx = 1
    title_fill = PatternFill("solid", fgColor="C6E0B4")
    header_fill = PatternFill("solid", fgColor="4472C4")

    for title, df_table in sections:
        ws.cell(row_idx, 1).value = title
        ws.cell(row_idx, 1).fill = title_fill
        ws.cell(row_idx, 1).font = Font(bold=True)
        row_idx += 1

        if df_table.empty:
            ws.cell(row_idx, 1).value = "Sin datos disponibles"
            row_idx += 3
            continue

        for col_idx, column in enumerate(df_table.columns, start=1):
            ws.cell(row_idx, col_idx).value = column
            ws.cell(row_idx, col_idx).fill = header_fill
            ws.cell(row_idx, col_idx).font = Font(bold=True, color="FFFFFF")
        row_idx += 1

        for record in df_table.itertuples(index=False):
            for col_idx, value in enumerate(record, start=1):
                cell = ws.cell(row_idx, col_idx)
                cell.value = value
                if str(df_table.columns[col_idx - 1]).endswith("%"):
                    cell.number_format = "0%"
            row_idx += 1
        row_idx += 2
    style_sheet(ws)


def build_excel_bytes(df, demo_cols, cfg1, cfg2, entities, aliases, normalizations, expected_raw_cols):
    run_1 = cfg1["tom"].strip() not in ["0", ""]
    run_2 = cfg2["tom"].strip() not in ["0", ""]

    if not run_1 and not run_2:
        raise ValueError("Debe configurar al menos un análisis válido.")

    wb = Workbook()
    wb.remove(wb.active)
    write_sections_sheet(wb, "Resumen Demográficos", demographic_sections(df, demo_cols))

    if run_1:
        output1, indicators1, _ = build_analysis(df, demo_cols, cfg1, entities, aliases, normalizations, expected_raw_cols)
        write_main_sheet(wb, cfg1["name"], output1)
        write_sections_sheet(wb, f"Resumen {cfg1['name']}", awareness_sections(indicators1, df, demo_cols, cfg1["name"], entities))
        write_sections_sheet(wb, f"Deptos {cfg1['name']}", department_sections(indicators1, df, demo_cols, entities))

    if run_2:
        output2, indicators2, _ = build_analysis(df, demo_cols, cfg2, entities, aliases, normalizations, expected_raw_cols)
        write_main_sheet(wb, cfg2["name"], output2)
        write_sections_sheet(wb, f"Resumen {cfg2['name']}", awareness_sections(indicators2, df, demo_cols, cfg2["name"], entities))
        write_sections_sheet(wb, f"Deptos {cfg2['name']}", department_sections(indicators2, df, demo_cols, entities))

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


# ============================================================
# INTERFAZ STREAMLIT
# ============================================================

st.set_page_config(page_title="Awareness Flexible", layout="wide")
st.title("Generador flexible de Awareness / Recordación")

with st.sidebar:
    st.header("Configuración general")
    mode = st.radio("Tipo de estudio", ["Bancos", "Conglomerados Financieros", "Personalizado"], index=0)
    uploaded_file = st.file_uploader("Archivo base .xlsx o .csv", type=["xlsx", "csv"])
    max_rows = st.number_input("Filas a evaluar (0 = todas)", min_value=0, value=0, step=50)
    
    # 🔴 SOLUCIÓN CLAVE: Ahora para Conglomerados, el campo por defecto es 0 (Sin validación estricta)
    default_expected_cols = 17 if mode == "Bancos" else 0
    expected_raw_cols = st.number_input("Columnas crudas esperadas (0 = no validar)", min_value=0, value=default_expected_cols, step=1)

# Asignación dinámica de diccionarios y listas según la selección
if mode == "Bancos":
    default_entities = BANKS
    default_aliases = BANK_ALIASES
    default_normalizations = BANK_NORMALIZATIONS
    def_sexo, def_edad, def_depto, def_estrato, def_ingreso = "F1 ", "F2 ", "F4 ", "F3 ", "0"
    t1_name, t1_tom, t1_esp, t1_ayu = "AWA PUB", "0", "0", "0"
    t2_name, t2_tom, t2_esp, t2_ayu = "AWA Marca", "P1", "P1A.1\nP1A.2\nP1A.3", "P2.1\nP2.2\nP2.3"
elif mode == "Conglomerados Financieros":
    default_entities = CONGLOMERATES
    default_aliases = CONGLOMERATE_ALIASES
    default_normalizations = []
    def_sexo, def_edad, def_depto, def_estrato, def_ingreso = "F1", "F2a", "F4", "F3", "F5"
    t1_name, t1_tom, t1_esp, t1_ayu = "Desactivado", "0", "0", "0"
    t2_name, t2_tom, t2_esp, t2_ayu = "Conglomerados AWA", "P1 -", "P1A -", "P2-P2."
else:
    default_entities = ["Entidad 1"]
    default_aliases = {"Entidad 1": ["entidad 1"]}
    default_normalizations = []
    def_sexo, def_edad, def_depto, def_estrato, def_ingreso = "", "", "", "", ""
    t1_name, t1_tom, t1_esp, t1_ayu = "Análisis 1", "0", "0", "0"
    t2_name, t2_tom, t2_esp, t2_ayu = "Análisis 2", "0", "0", "0"

with st.expander("1. Entidades, alias y normalizaciones", expanded=(mode == "Personalizado")):
    entities_text = st.text_area("Entidades a evaluar, una por línea", "\n".join(default_entities), height=200)
    aliases_text = st.text_area("Alias / condiciones en JSON", json.dumps(default_aliases, ensure_ascii=False, indent=2), height=220)
    normalizations_text = st.text_area("Normalizaciones opcionales en JSON", json.dumps(default_normalizations, ensure_ascii=False, indent=2), height=120)

try:
    aliases = json.loads(aliases_text)
    normalizations = json.loads(normalizations_text)
except Exception as exc:
    st.error(f"Error en estructuras JSON: {exc}")
    st.stop()

entities = [line.strip() for line in entities_text.splitlines() if line.strip()]

with st.expander("2. Preguntas y demográficos", expanded=True):
    col1, col2, col3 = st.columns(3)
    with col1:
        sexo_prefix = st.text_input("Sexo", def_sexo)
        edad_prefix = st.text_input("Edad", def_edad)
        departamento_prefix = st.text_input("Departamento / Ciudad", def_depto)
        estrato_prefix = st.text_input("Estrato", def_estrato)
        ingreso_prefix = st.text_input("Ingreso / Rangos", def_ingreso)
    with col2:
        a1_name = st.text_input("Nombre análisis 1", t1_name)
        a1_tom = st.text_input("TOM análisis 1", t1_tom)
        a1_esp = st.text_area("Espontáneo análisis 1", t1_esp, height=100)
        a1_ayu = st.text_area("Ayudado análisis 1", t1_ayu, height=100)
    with col3:
        a2_name = st.text_input("Nombre análisis 2", t2_name)
        a2_tom = st.text_input("TOM análisis 2", t2_tom)
        a2_esp = st.text_area("Espontáneo análisis 2", t2_esp, height=100)
        a2_ayu = st.text_area("Ayudado análisis 2", t2_ayu, height=100)

if uploaded_file is None:
    st.info("Por favor carga el archivo de datos para iniciar.")
    st.stop()

try:
    if uploaded_file.name.endswith('.csv'):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file, engine="openpyxl")
except Exception as exc:
    st.error(f"Error al leer el archivo: {exc}")
    st.stop()

if max_rows and int(max_rows) > 0:
    df = df.head(int(max_rows)).copy()

demo_cols = {
    "sexo": find_col(df, sexo_prefix),
    "edad": find_col(df, edad_prefix),
    "departamento": find_col(df, departamento_prefix),
    "estrato": find_col(df, estrato_prefix),
    "ingreso": find_col(df, ingreso_prefix),
}

cfg1 = {"name": a1_name, "tom": a1_tom, "esp": a1_esp, "ayu": a1_ayu}
cfg2 = {"name": a2_name, "tom": a2_tom, "esp": a2_esp, "ayu": a2_ayu}

st.subheader("3. Validación de Columnas")
validation_rows = []
for key, value in demo_cols.items():
    validation_rows.append({"Tipo": "Demográfico", "Campo": key, "Columna detectada": str(value)})
for label, cfg in [("Análisis 1", cfg1), ("Análisis 2", cfg2)]:
    if cfg["tom"].strip() not in ["0", ""]:
        validation_rows.append({"Tipo": label, "Campo": "TOM", "Columna detectada": str(find_col(df, cfg["tom"]))})
        # Verás cómo ahora el Espontáneo y Ayudado agrupan múltiples columnas separadas por comas.
        validation_rows.append({"Tipo": label, "Campo": "Espontáneo", "Columna detectada": ", ".join(find_cols(df, cfg["esp"]))})
        validation_rows.append({"Tipo": label, "Campo": "Ayudado", "Columna detectada": ", ".join(find_cols(df, cfg["ayu"]))})

st.dataframe(pd.DataFrame(validation_rows), use_container_width=True)

st.subheader("4. Procesar y Descargar")
if st.button("Generar Reporte Excel", type="primary"):
    try:
        result = build_excel_bytes(
            df=df, demo_cols=demo_cols, cfg1=cfg1, cfg2=cfg2,
            entities=entities, aliases=aliases, normalizations=normalizations,
            expected_raw_cols=int(expected_raw_cols)
        )
        st.success("¡Cálculo e indicadores generados exitosamente!")
        st.download_button(
            label="📥 Descargar Reporte de Awareness",
            data=result,
            file_name="Reporte_Awareness_Estructural.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as exc:
        st.exception(exc)
