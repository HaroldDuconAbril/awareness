# -*- coding: utf-8 -*-

"""
Generador Flexible de Awareness / Recordación - Streamlit

Funciona para:
- Bancos
- Conglomerados financieros
- Marcas
- Empresas
- Productos
- Países
- Cualquier entidad configurable por el usuario

Ejecutar localmente:
    streamlit run streamlit_awareness_generico.py
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
# CONFIGURACIÓN PREDEFINIDA PARA BANCOS
# ============================================================

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
    "Bancolombia": [
        "bancolombia",
        "bamcolombia",
        "bancoloombia",
        "bsncolombia",
        "bancooombia",
    ],
    "Davivienda": ["davivienda"],
    "Banco de Bogotá": [
        "banco de bogota",
        "banco de bogotá",
        "banco bogota",
        "banco bogotá",
        "bancobogota",
    ],
    "BBVA": ["bbva", "bbvva", "bvva", "bbwa", "bva", "bvvwa"],
    "Banco Caja Social": [
        "banco caja social",
        "caja social",
        "cajas social",
        "caja sosial",
    ],
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
    {
        "pattern": r"(?i)Bamcolombia|Bancoloombia|Bsncolombia|Bancooombia",
        "replacement": "Bancolombia",
    },
]


# ============================================================
# CONFIGURACIÓN PREDEFINIDA PARA CONGLOMERADOS FINANCIEROS
# ============================================================

CONGLOMERADOS = [
    "Conglomerado BBVA",
    "Grupo cooperativo Coomeva",
    "Fundación Grupo Social",
    "Grupo Bolivar",
    "Conglomerado financiero Sura-Bancolombia",
    "Grupo Aval",
    "BPN Paribas",
    "BTG Pactual",
    "GNB Sudameris",
    "Conglomerado Credicorp capital",
    "Conglomerado Skandia",
]

CONGLOMERADOS_ALIASES = {
    "Conglomerado BBVA": [
        "conglomerado bbva",
        "grupo bbva",
        "bbva",
        "bbvva",
        "bvva",
        "bbwa",
        "bva",
    ],
    "Grupo cooperativo Coomeva": [
        "grupo cooperativo coomeva",
        "coomeva",
        "grupo coomeva",
    ],
    "Fundación Grupo Social": [
        "fundacion grupo social",
        "fundación grupo social",
        "grupo social",
        "fundacion social",
        "fundación social",
        "caja social",
        "grupo caja social",
        "banco caja social",
    ],
    "Grupo Bolivar": [
        "grupo bolivar",
        "grupo bolívar",
        "bolivar",
        "bolívar",
        "davivienda",
        "grupo bolivar davivienda",
        "grupo bolívar davivienda",
    ],
    "Conglomerado financiero Sura-Bancolombia": [
        "conglomerado financiero sura bancolombia",
        "conglomerado financiero sura-bancolombia",
        "sura bancolombia",
        "sura-bancolombia",
        "grupo sura bancolombia",
        "grupo sura",
        "sura",
        "bancolombia",
        "grupo bancolombia",
        "banco bancolombia",
        "bamcolombia",
        "bancoloombia",
        "bsncolombia",
        "bancooombia",
    ],
    "Grupo Aval": [
        "grupo aval",
        "aval",
        "bancos aval",
        "grupó aval",
        "grupo av villas",
        "av villas",
        "avvillas",
        "ave villas",
        "banco de bogota",
        "banco de bogotá",
        "banco bogota",
        "bancobogota",
        "banco popular",
        "banco de occidente",
    ],
    "BPN Paribas": [
        "bpn paribas",
        "bnp paribas",
        "paribas",
    ],
    "BTG Pactual": [
        "btg pactual",
        "btg",
    ],
    "GNB Sudameris": [
        "gnb sudameris",
        "sudameris",
        "gnb",
    ],
    "Conglomerado Credicorp capital": [
        "conglomerado credicorp capital",
        "credicorp capital",
        "credicorp",
    ],
    "Conglomerado Skandia": [
        "conglomerado skandia",
        "skandia",
    ],
}

CONGLOMERADOS_NORMALIZATIONS = [
    {"pattern": r"(?i)Grupo\s+Bol[ií]var", "replacement": "Grupo Bolivar"},
    {
        "pattern": r"(?i)Fundacion\s+Grupo\s+Social",
        "replacement": "Fundación Grupo Social",
    },
    {
        "pattern": r"(?i)Fundacion\s+Social",
        "replacement": "Fundación Grupo Social",
    },
    {
        "pattern": r"(?i)Sura\s*[- ]\s*Bancolombia",
        "replacement": "Sura-Bancolombia",
    },
    {
        "pattern": r"(?i)Bamcolombia|Bancoloombia|Bsncolombia|Bancooombia",
        "replacement": "Bancolombia",
    },
    {"pattern": r"(?i)BPN\s+Paribas", "replacement": "BPN Paribas"},
    {"pattern": r"(?i)BNP\s+Paribas", "replacement": "BPN Paribas"},
]


# ============================================================
# NEGATIVOS
# ============================================================

NEGATIVOS = {
    "",
    "0",
    "no",
    "nan",
    "none",
    "false",
    "ninguno",
    "ninguna",
    "ningun",
    "ningún",
    "no se",
    "no sé",
    "nose",
    "no recuerdo",
    "no conozco",
    "no conozco mas",
    "no conozco más",
    "ningun otro",
    "ningún otro",
    "ninguna mas",
    "ninguna más",
    "no aplica",
    "ni idea",
    "n/a",
    "na",
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
    """Excel limita nombres de hoja a 31 caracteres y no permite algunos símbolos."""
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

        if not alias_norm:
            continue

        pattern = r"(^|\s)" + re.escape(alias_norm) + r"(\s|$)"

        if re.search(pattern, text):
            return True

    return False


# ============================================================
# DETECCIÓN DE COLUMNAS
# ============================================================

def find_col(df: pd.DataFrame, prefix: str) -> Optional"""Encuentra la primera columna cuyo encabezado empieza por el prefijo indicado."""
    prefix_norm = norm(prefix)

    if not prefix_norm or prefix_norm == "0":
        return None

    for column in df.columns:
        if norm(column).startswith(prefix_norm):
            return column

    return None


def prefixes_to_list(text: str) -> List"""Convierte un textarea en lista de prefijos, uno por línea."""
    return [line.strip() for line in str(text).splitlines() if line.strip()]


def find_cols(df: pd.DataFrame, prefixes_text: str) -> List"""
    Encuentra columnas para varios prefijos.

    Importante:
    Esta función trae TODAS las columnas que empiezan por cada prefijo.
    Así P2-P2 detecta todas las columnas ayudadas de la nueva base.
    """
    detected = []
    seen = set()

    for prefix in prefixes_to_list(prefixes_text):
        prefix_norm = norm(prefix)

        if not prefix_norm or prefix_norm == "0":
            continue

        for column in df.columns:
            if norm(column).startswith(prefix_norm):
                if column not in seen:
                    detected.append(column)
                    seen.add(column)

    return detected


def entity_from_aided_col(column_name: str, entities: List[str]) -> Optional"""
    Detecta a qué entidad corresponde una columna ayudada según el texto del encabezado.

    Ejemplo:
    P2-P2. Del siguiente listado ... - Conglomerado BBVA
    """
    original = str(column_name).replace("\n", " ")
    text = norm(original)

    parts = re.split(r"\s-\s", original)
    last_part = norm(parts[-1]) if parts else text

    for entity in entities:
        entity_norm = norm(entity)

        if entity_norm in text or entity_norm in last_part:
            return entity

    if "Conglomerado BBVA" in entities and "bbva" in text:
        return "Conglomerado BBVA"

    if "Grupo cooperativo Coomeva" in entities and "coomeva" in text:
        return "Grupo cooperativo Coomeva"

    if "Fundación Grupo Social" in entities and (
        "fundacion grupo social" in text
        or "fundacion social" in text
        or "grupo social" in text
        or "caja social" in text
    ):
        return "Fundación Grupo Social"

    if "Grupo Bolivar" in entities and "bolivar" in text:
        return "Grupo Bolivar"

    if "Conglomerado financiero Sura-Bancolombia" in entities and (
        "sura bancolombia" in text
        or "sura bancolombia" in last_part
        or "bancolombia" in text
    ):
        return "Conglomerado financiero Sura-Bancolombia"

    if "Grupo Aval" in entities and "grupo aval" in text:
        return "Grupo Aval"

    if "BPN Paribas" in entities and (
        "bpn paribas" in text or "bnp paribas" in text
    ):
        return "BPN Paribas"

    if "BTG Pactual" in entities and "btg pactual" in text:
        return "BTG Pactual"

    if "GNB Sudameris" in entities and "gnb sudameris" in text:
        return "GNB Sudameris"

    if "Conglomerado Credicorp capital" in entities and "credicorp" in text:
        return "Conglomerado Credicorp capital"

    if "Conglomerado Skandia" in entities and "skandia" in text:
        return "Conglomerado Skandia"

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
    raw_cols = list(dict.fromkeys(raw_cols))

    if expected_raw_cols and int(expected_raw_cols) > 0 and len(raw_cols) != int(expected_raw_cols):
        raise ValueError(
            f"{cfg['name']}: se esperaban {expected_raw_cols} columnas crudas "
            f"y se detectaron {len(raw_cols)}. Columnas detectadas: {raw_cols}"
        )

    if not tom_col:
        raise ValueError(f"{cfg['name']}: no se detectó la columna TOM.")

    if not esp_cols:
        raise ValueError(f"{cfg['name']}: no se detectaron columnas espontáneas.")

    if not ayud_cols:
        raise ValueError(f"{cfg['name']}: no se detectaron columnas ayudadas.")

    raw_df = df[raw_cols].apply(
        lambda col: col.map(lambda value: apply_normalizations(value, normalizations))
    )

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
            espontaneo = espontaneo | raw_df[col].apply(
                lambda value: contains_entity(value, entity, aliases)
            )

        aided_col = aided_map.get(entity)

        if aided_col:
            def check_aided_value(value, ent, als):
                if pd.isna(value):
                    return False

                v_norm = norm(value)

                if v_norm in NEGATIVOS:
                    return False

                if v_norm in ["1", "si", "sí", "x", "seleccionado", "seleccionada", "true"]:
                    return True

                return contains_entity(value, ent, als)

            ayudado = raw_df[aided_col].apply(
                lambda value: check_aided_value(value, entity, aliases)
            )

        else:
            ayudado = pd.Series(False, index=df.index)

            for col in ayud_cols:
                ayudado = ayudado | raw_df[col].apply(
                    lambda value: contains_entity(value, entity, aliases)
                )

        indicators[(entity, "TOM")] = tom.astype(int)
        indicators[(entity, "Espontaneo")] = ((~tom) & espontaneo).astype(int)
        indicators[(entity, "Ayudado")] = ((~tom) & (~espontaneo) & ayudado).astype(int)

        output_df[(entity, "TOM")] = indicators[(entity, "TOM")]
        output_df[(entity, "Espontaneo")] = indicators[(entity, "Espontaneo")]
        output_df[(entity, "Ayudado")] = indicators[(entity, "Ayudado")]

    return output_df, pd.DataFrame(indicators), raw_cols


# ============================================================
# TABLAS RESUMEN
# ============================================================

def category_table(
    indicators: pd.DataFrame,
    df: pd.DataFrame,
    category_col: Optional[str],
    category_value,
    entities: List[str],
) -> pd.DataFrame:

    if category_col is None:
        mask = pd.Series(True, index=df.index)
    else:
        mask = df[category_col].fillna("Sin dato") == category_value

    base = int(mask.sum())
    rows = []

    for entity in entities:
        tom = int(indicators[(entity, "TOM")][mask].sum())
        esp = int(indicators[(entity, "Espontaneo")][mask].sum())
        ayu = int(indicators[(entity, "Ayudado")][mask].sum())
        awa = tom + esp + ayu

        rows.append(
            {
                "Entidad": entity,
                "Base": base,
                "TOM": tom,
                "TOM %": tom / base if base else 0,
                "Espontaneo": esp,
                "Espontaneo %": esp / base if base else 0,
                "Ayudado": ayu,
                "Ayudado %": ayu / base if base else 0,
                "AWA": awa,
                "AWA %": awa / base if base else 0,
            }
        )

    return pd.DataFrame(rows)


def awareness_sections(
    indicators: pd.DataFrame,
    df: pd.DataFrame,
    demo_cols: Dict,
    analysis_name: str,
    entities: List[str],
):
    sections = [
        (
            f"Resumen general {analysis_name}",
            category_table(indicators, df, None, None, entities),
        )
    ]

    sexo_col = demo_cols.get("sexo")

    if sexo_col:
        for value in pd.Series(df[sexo_col].fillna("Sin dato")).drop_duplicates():
            sections.append(
                (
                    f"Detalle por sexo - {value}",
                    category_table(indicators, df, sexo_col, value, entities),
                )
            )

    estrato_col = demo_cols.get("estrato")

    if estrato_col:
        for value in pd.Series(df[estrato_col].fillna("Sin dato")).drop_duplicates():
            sections.append(
                (
                    f"Detalle por estrato - {value}",
                    category_table(indicators, df, estrato_col, value, entities),
                )
            )

    return sections


def department_sections(
    indicators: pd.DataFrame,
    df: pd.DataFrame,
    demo_cols: Dict,
    entities: List[str],
):
    dept_col = demo_cols.get("departamento")

    if not dept_col:
        return [("Sin departamento / ciudad detectado", pd.DataFrame())]

    sections = []

    for value in pd.Series(df[dept_col].fillna("Sin dato")).drop_duplicates():
        sections.append(
            (
                str(value),
                category_table(indicators, df, dept_col, value, entities),
            )
        )

    return sections


def demographic_sections(df: pd.DataFrame, demo_cols: Dict):
    sections = [
        (
            "Base",
            pd.DataFrame(
                {
                    "Indicador": ["Total encuestas procesadas"],
                    "Valor": [len(df)],
                }
            ),
        )
    ]

    for key, label in [
        ("sexo", "Sexo"),
        ("edad", "Edad"),
        ("departamento", "Ciudad / Departamento"),
        ("estrato", "Estrato"),
        ("ingreso", "Ingreso"),
    ]:
        col = demo_cols.get(key)

        if col:
            table = (
                df[col]
                .fillna("Sin dato")
                .value_counts()
                .rename_axis(label)
                .reset_index(name="Cantidad")
            )

            table["%"] = table["Cantidad"] / len(df) if len(df) else 0

            sections.append((f"Demográfico - {label}", table))

    if demo_cols.get("departamento") and demo_cols.get("sexo"):
        cross = pd.crosstab(
            df[demo_cols["departamento"]].fillna("Sin dato"),
            df[demo_cols["sexo"]].fillna("Sin dato"),
        )

        cross["Total"] = cross.sum(axis=1)

        cross = cross.reset_index().rename(
            columns={demo_cols["departamento"]: "Ciudad / Departamento"}
        )

        sections.append(("Participación Sexo por Ciudad / Departamento", cross))

    return sections


# ============================================================
# ESCRITURA EXCEL
# ============================================================

def style_sheet(ws):
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for row in ws.iter_rows():
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True,
            )

    for col_idx in range(1, ws.max_column + 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = 20


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


def build_excel_bytes(
    df,
    demo_cols,
    cfg1,
    cfg2,
    entities,
    aliases,
    normalizations,
    expected_raw_cols,
):
    run_1 = cfg1["tom"].strip() not in ["0", ""]
    run_2 = cfg2["tom"].strip() not in ["0", ""]

    if not run_1 and not run_2:
        raise ValueError(
            "Debe configurar al menos un análisis válido: TOM diferente de 0 o vacío."
        )

    wb = Workbook()
    wb.remove(wb.active)

    write_sections_sheet(
        wb,
        "Resumen Demográficos",
        demographic_sections(df, demo_cols),
    )

    if run_1:
        output1, indicators1, _ = build_analysis(
            df,
            demo_cols,
            cfg1,
            entities,
            aliases,
            normalizations,
            expected_raw_cols,
        )

        write_main_sheet(wb, cfg1["name"], output1)

        write_sections_sheet(
            wb,
            f"Resumen {cfg1['name']}",
            awareness_sections(indicators1, df, demo_cols, cfg1["name"], entities),
        )

        write_sections_sheet(
            wb,
            f"Deptos {cfg1['name']}",
            department_sections(indicators1, df, demo_cols, entities),
        )

    if run_2:
        output2, indicators2, _ = build_analysis(
            df,
            demo_cols,
            cfg2,
            entities,
            aliases,
            normalizations,
            expected_raw_cols,
        )

        write_main_sheet(wb, cfg2["name"], output2)

        write_sections_sheet(
            wb,
            f"Resumen {cfg2['name']}",
            awareness_sections(indicators2, df, demo_cols, cfg2["name"], entities),
        )

        write_sections_sheet(
            wb,
            f"Deptos {cfg2['name']}",
            department_sections(indicators2, df, demo_cols, entities),
        )

    output = io.BytesIO()
    wb.save(output)

    return output.getvalue()


# ============================================================
# INTERFAZ STREAMLIT
# ============================================================

st.set_page_config(page_title="Awareness Flexible", layout="wide")

st.title("Generador flexible de Awareness / Recordación")
st.caption(
    "Funciona para bancos, conglomerados financieros, marcas, empresas, productos, países u otras entidades."
)


with st.sidebar:
    st.header("Configuración general")

    mode = st.radio(
        "Tipo de estudio",
        ["Bancos", "Conglomerados financieros", "Personalizado"],
        index=1,
    )

    uploaded_file = st.file_uploader("Archivo base .xlsx", type=["xlsx"])

    max_rows = st.number_input(
        "Filas a evaluar (0 = todas)",
        min_value=0,
        value=0,
        step=50,
    )

    default_expected_cols = 0 if mode == "Conglomerados financieros" else 17

    expected_raw_cols = st.number_input(
        "Columnas crudas esperadas (0 = no validar)",
        min_value=0,
        value=default_expected_cols,
        step=1,
    )


# ============================================================
# DEFAULTS SEGÚN MODO
# ============================================================

if mode == "Bancos":
    default_entities = BANKS
    default_aliases = BANK_ALIASES
    default_normalizations = BANK_NORMALIZATIONS

    default_sexo = "F1 "
    default_edad = "F2 "
    default_departamento = "F4 "
    default_estrato = "F3 "
    default_ingreso = "0"

    default_a1_name = "AWA PUB"
    default_a1_tom = "0"
    default_a1_esp = "0"
    default_a1_ayu = "0"

    default_a2_name = "AWA Marca"
    default_a2_tom = "P1"
    default_a2_esp = "\n".join([f"P1A.{i}" for i in range(1, 11)])
    default_a2_ayu = "\n".join([f"P2.{i}" for i in range(1, 6)])

elif mode == "Conglomerados financieros":
    default_entities = CONGLOMERADOS
    default_aliases = CONGLOMERADOS_ALIASES
    default_normalizations = CONGLOMERADOS_NORMALIZATIONS

    default_sexo = "F1 "
    default_edad = "F2a "
    default_departamento = "F4 "
    default_estrato = "F3 "
    default_ingreso = "F5 "

    default_a1_name = "AWA PUB"
    default_a1_tom = "0"
    default_a1_esp = "0"
    default_a1_ayu = "0"

    default_a2_name = "AWA Conglomerados"
    default_a2_tom = "P1"
    default_a2_esp = "P1A"
    default_a2_ayu = "P2-P2"

else:
    default_entities = ["Entidad 1", "Entidad 2", "Entidad 3"]

    default_aliases = {
        "Entidad 1": ["entidad 1"],
        "Entidad 2": ["entidad 2"],
        "Entidad 3": ["entidad 3"],
    }

    default_normalizations = []

    default_sexo = "F1 "
    default_edad = "F2 "
    default_departamento = "F4 "
    default_estrato = "F3 "
    default_ingreso = "0"

    default_a1_name = "Awareness 1"
    default_a1_tom = "0"
    default_a1_esp = "0"
    default_a1_ayu = "0"

    default_a2_name = "Awareness 2"
    default_a2_tom = "P1"
    default_a2_esp = "P1A"
    default_a2_ayu = "P2"


# ============================================================
# CONFIGURACIÓN DE ENTIDADES, ALIAS Y NORMALIZACIONES
# ============================================================

with st.expander(
    "1. Entidades, alias y normalizaciones",
    expanded=(mode == "Personalizado"),
):
    entities_text = st.text_area(
        "Entidades a evaluar, una por línea",
        "\n".join(default_entities),
        height=220,
    )

    aliases_text = st.text_area(
        "Alias / condiciones en JSON",
        json.dumps(default_aliases, ensure_ascii=False, indent=2),
        height=280,
    )

    normalizations_text = st.text_area(
        "Normalizaciones opcionales en JSON",
        json.dumps(default_normalizations, ensure_ascii=False, indent=2),
        height=180,
    )


try:
    aliases = json.loads(aliases_text)
except Exception as exc:
    st.error(f"El JSON de alias no es válido: {exc}")
    st.stop()


try:
    normalizations = json.loads(normalizations_text)
except Exception as exc:
    st.error(f"El JSON de normalizaciones no es válido: {exc}")
    st.stop()


entities = [line.strip() for line in entities_text.splitlines() if line.strip()]


# ============================================================
# CONFIGURACIÓN DE PREGUNTAS Y DEMOGRÁFICOS
# ============================================================

with st.expander("2. Preguntas y demográficos", expanded=True):
    col1, col2, col3 = st.columns(3)

    with col1:
        sexo_prefix = st.text_input("Sexo", default_sexo)
        edad_prefix = st.text_input("Edad", default_edad)
        departamento_prefix = st.text_input(
            "Ciudad / Departamento / Región / Zona",
            default_departamento,
        )
        estrato_prefix = st.text_input("Estrato / NSE / Segmento", default_estrato)
        ingreso_prefix = st.text_input("Ingreso / Otra variable", default_ingreso)

    with col2:
        a1_name = st.text_input("Nombre análisis 1", default_a1_name)
        a1_tom = st.text_input("TOM análisis 1", default_a1_tom)
        a1_esp = st.text_area(
            "Espontáneo análisis 1, uno por línea",
            default_a1_esp,
        )
        a1_ayu = st.text_area(
            "Ayudado análisis 1, uno por línea",
            default_a1_ayu,
        )

    with col3:
        a2_name = st.text_input("Nombre análisis 2", default_a2_name)
        a2_tom = st.text_input("TOM análisis 2", default_a2_tom)
        a2_esp = st.text_area(
            "Espontáneo análisis 2, uno por línea",
            default_a2_esp,
        )
        a2_ayu = st.text_area(
            "Ayudado análisis 2, uno por línea",
            default_a2_ayu,
        )


# ============================================================
# LECTURA DE ARCHIVO
# ============================================================

if uploaded_file is None:
    st.info("Carga un archivo Excel para iniciar.")
    st.stop()


try:
    df = pd.read_excel(uploaded_file, engine="openpyxl")
except Exception as exc:
    st.error(f"No se pudo leer el archivo Excel: {exc}")
    st.stop()


if max_rows and int(max_rows) > 0:
    df = df.head(int(max_rows)).copy()


# ============================================================
# DETECCIÓN DE DEMOGRÁFICOS Y CONFIGURACIÓN
# ============================================================

demo_cols = {
    "sexo": find_col(df, sexo_prefix),
    "edad": find_col(df, edad_prefix),
    "departamento": find_col(df, departamento_prefix),
    "estrato": find_col(df, estrato_prefix),
    "ingreso": find_col(df, ingreso_prefix),
}

cfg1 = {
    "name": a1_name,
    "tom": a1_tom,
    "esp": a1_esp,
    "ayu": a1_ayu,
}

cfg2 = {
    "name": a2_name,
    "tom": a2_tom,
    "esp": a2_esp,
    "ayu": a2_ayu,
}


# ============================================================
# VALIDACIÓN
# ============================================================

st.subheader("3. Validación")

validation_rows = []

for key, value in demo_cols.items():
    validation_rows.append(
        {
            "Tipo": "Demográfico",
            "Campo": key,
            "Columna detectada": str(value),
        }
    )

for label, cfg in [("Análisis 1", cfg1), ("Análisis 2", cfg2)]:
    validation_rows.append(
        {
            "Tipo": label,
            "Campo": "TOM",
            "Columna detectada": str(find_col(df, cfg["tom"])),
        }
    )

    validation_rows.append(
        {
            "Tipo": label,
            "Campo": "Espontáneo",
            "Columna detectada": ", ".join(find_cols(df, cfg["esp"])),
        }
    )

    validation_rows.append(
        {
            "Tipo": label,
            "Campo": "Ayudado",
            "Columna detectada": ", ".join(find_cols(df, cfg["ayu"])),
        }
    )

st.dataframe(pd.DataFrame(validation_rows), use_container_width=True)


with st.expander(
    "Validación avanzada: mapeo de columnas ayudadas a entidades",
    expanded=False,
):
    ayud_cols_preview = find_cols(df, cfg2["ayu"])
    map_rows = []

    for col in ayud_cols_preview:
        map_rows.append(
            {
                "Columna ayudada": col,
                "Entidad detectada": entity_from_aided_col(col, entities),
            }
        )

    if map_rows:
        st.dataframe(pd.DataFrame(map_rows), use_container_width=True)
    else:
        st.info("No se detectaron columnas ayudadas para el análisis 2.")


# ============================================================
# GENERACIÓN DE ARCHIVO
# ============================================================

st.subheader("4. Generar archivo")

if st.button("Generar Excel", type="primary"):
    try:
        result = build_excel_bytes(
            df=df,
            demo_cols=demo_cols,
            cfg1=cfg1,
            cfg2=cfg2,
            entities=entities,
            aliases=aliases,
            normalizations=normalizations,
            expected_raw_cols=int(expected_raw_cols),
        )

        st.success("Archivo generado correctamente.")

        st.download_button(
            label="Descargar resultado",
            data=result,
            file_name="Resultado_Awareness_Flexible.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as exc:
        st.exception(exc)
