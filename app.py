import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime, date, timedelta
import re

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation


APP_VERSION = "v2026-01-06_ROBUSTO_EXCEL_FORMULAS"

st.set_page_config(page_title="Ausentismo e Incidencias Operativas", layout="wide")
st.title("Plataforma de Gestión de Ausentismo e Incidencias Operativas")
st.caption(APP_VERSION)

# =========================
# Paleta / estilo (Cabify-ish)
# =========================
CABIFY_HEADER = "362065"   # morado
CABIFY_ACCENT = "E83C96"   # contraste
GRID = "D9D9D9"
WHITE = "FFFFFF"

thin = Side(style="thin", color=GRID)
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

def style_header_row(ws, row=1, fill_hex=CABIFY_HEADER):
    fill = PatternFill("solid", fgColor=fill_hex)
    font = Font(color=WHITE, bold=True)
    for cell in ws[row]:
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER

def autosize_columns(ws, max_width=45):
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for c in col:
            v = c.value
            if v is None:
                continue
            s = str(v)
            max_len = max(max_len, len(s))
        ws.column_dimensions[col_letter].width = min(max_len + 2, max_width)

# =========================
# Helpers de parsing
# =========================
def normalize_rut(x):
    if pd.isna(x):
        return ""
    return str(x).strip().upper().replace(".", "").replace(" ", "")

def try_parse_date_any(x):
    if pd.isna(x):
        return pd.NaT
    return pd.to_datetime(x, errors="coerce", dayfirst=True)

def _clean_cell(x):
    if pd.isna(x):
        return ""
    return str(x).replace("\u00A0", " ").strip()

def _norm_colname(s):
    s = "" if s is None else str(s)
    s = s.replace("\u00A0", " ").lower().strip()
    for ch in [" ", ".", "-", "_", "\n", "\t", "\r", ":", ";", ",", "(", ")"]:
        s = s.replace(ch, "")
    return s

def find_col(df, candidates):
    norm = {_norm_colname(c): c for c in df.columns}
    for c in candidates:
        k = _norm_colname(c)
        if k in norm:
            return norm[k]
    return None

def read_raw(file, sheet=0):
    name = getattr(file, "name", "").lower()
    if name.endswith(".xls"):
        return pd.read_excel(file, sheet_name=sheet, header=None, engine="xlrd")
    return pd.read_excel(file, sheet_name=sheet, header=None, engine="openpyxl")

def ffill_row(values):
    out, last = [], ""
    for v in values:
        s = _clean_cell(v)
        if s:
            last = s
            out.append(s)
        else:
            out.append(last)
    return out

# Asistencia con encabezado 2 filas tipo Entrada/Salida + Fecha/Hora
def read_asistencia(file):
    raw = read_raw(file)
    h1 = ffill_row(raw.iloc[0].tolist())
    h2 = [_clean_cell(v) for v in raw.iloc[1].tolist()]

    # a veces hay columna 0 vacía; ajusta si hace falta
    start = 1 if (h1 and h1[0] == "" and h2 and h2[0] == "") else 0
    h1, h2 = h1[start:], h2[start:]

    cols = []
    for g, s in zip(h1, h2):
        gn, sn = _norm_colname(g), _norm_colname(s)

        if gn == "entrada" and sn == "fecha":
            cols.append("Fecha Entrada")
        elif gn == "entrada" and sn == "hora":
            cols.append("Hora Entrada")
        elif gn == "salida" and sn == "fecha":
            cols.append("Fecha Salida")
        elif gn == "salida" and sn == "hora":
            cols.append("Hora Salida")
        else:
            # fallback razonable
            if s and g:
                cols.append(f"{g} {s}".strip())
            elif s:
                cols.append(s)
            else:
                cols.append(g if g else "")

    cols = [_clean_cell(c) if c else f"COL_{i}" for i, c in enumerate(cols)]
    df = raw.iloc[2:, start:].copy()
    df.columns = cols
    df = df.dropna(how="all")
    return df, raw

def read_inasist(file):
    raw = read_raw(file)
    header = None
    for i in range(min(120, len(raw))):
        row = [_norm_colname(v) for v in raw.iloc[i].tolist()]
        if "rut" in row and ("dia" in row or "día" in row):
            header = i
            break
    if header is None:
        raise RuntimeError("No pude detectar encabezado en Inasistencias (no veo RUT + Día).")

    cols = [_clean_cell(c) for c in raw.iloc[header].tolist()]
    df = raw.iloc[header + 1 :].copy()
    df.columns = cols
    df = df.dropna(how="all")
    return df, raw

def norm_recinto(x):
    if pd.isna(x):
        return "Sin Marca"
    s = str(x).strip()
    su = s.upper().replace("SÍ", "SI")

    if su in ["SI", "S"]:
        return "Sí"
    if su in ["NO", "N"]:
        return "No"
    if su in ["SIN COORDENADAS", "SINCOORDENADAS"]:
        return "No"
    if su == "":
        return "Sin Marca"
    return s

def parse_horario_to_hours(horario: str) -> float:
    """
    horario ejemplo: "7:00:00 - 15:00:00" o "20:00:00 - 07:00:00"
    retorna horas (float)
    """
    if not isinstance(horario, str):
        return np.nan
    m = re.split(r"\s*-\s*", horario.strip())
    if len(m) != 2:
        return np.nan
    try:
        t1 = pd.to_datetime(m[0].strip()).time()
        t2 = pd.to_datetime(m[1].strip()).time()
    except Exception:
        return np.nan

    dt1 = datetime(2000, 1, 1, t1.hour, t1.minute, t1.second)
    dt2 = datetime(2000, 1, 1, t2.hour, t2.minute, t2.second)
    if dt2 <= dt1:
        dt2 += timedelta(days=1)
    return (dt2 - dt1).total_seconds() / 3600.0

def safe_time_str(x):
    if pd.isna(x):
        return ""
    try:
        t = pd.to_datetime(x).time()
        return f"{t.hour:02d}:{t.minute:02d}"
    except Exception:
        s = str(x).strip()
        return s[:5] if len(s) >= 5 else s

# =========================
# UI - uploads + parámetros
# =========================
with st.sidebar:
    st.header("Inputs")
    f_asist = st.file_uploader("1) Asistencia (XLS/XLSX)", type=["xls", "xlsx"])
    f_inas = st.file_uploader("2) Inasistencias (XLS/XLSX)", type=["xls", "xlsx"])
    f_plan = st.file_uploader("3) Planificación de Turnos (CSV)", type=["csv"])
    f_cod = st.file_uploader("4) Codificación de Turno (CSV)", type=["csv"])

    st.divider()
    st.subheader("Filtros / reglas")
    only_area = st.text_input("Filtrar Área (opcional)", value="AEROPUERTO")
    umbral_diff_h = st.number_input("Diferencia mínima para incidencia (horas)", value=0.5, step=0.25)
    st.caption("Se usa para filtrar incidencias por diferencia (planificado vs trabajado).")

if not all([f_asist, f_inas, f_plan, f_cod]):
    st.info("Sube los 4 archivos para comenzar.")
    st.stop()

# =========================
# Load data
# =========================
df_asist, raw_asist = read_asistencia(f_asist)
df_inas, raw_inas = read_inasist(f_inas)

df_plan = pd.read_csv(f_plan)
df_cod = pd.read_csv(f_cod)

# Diagnóstico
with st.expander("Diagnóstico de lectura (útil si cambia el formato)", expanded=False):
    st.write("Asistencia - columnas:", list(df_asist.columns))
    st.dataframe(df_asist.head(5), use_container_width=True)
    st.write("Inasistencias - columnas:", list(df_inas.columns))
    st.dataframe(df_inas.head(5), use_container_width=True)
    st.write("Planificación - columnas:", list(df_plan.columns))
    st.dataframe(df_plan.head(5), use_container_width=True)
    st.write("Codificación - columnas:", list(df_cod.columns))
    st.dataframe(df_cod.head(5), use_container_width=True)

# =========================
# Normalizaciones + filtros
# =========================
# Asistencia: RUT + Fechas
c_rut_a = find_col(df_asist, ["RUT"])
if not c_rut_a:
    st.error("Asistencia: no encontré columna RUT.")
    st.stop()

df_asist["RUT_norm"] = df_asist[c_rut_a].apply(normalize_rut)

# Dentro recinto
c_rec_in = find_col(df_asist, ["Dentro del Recinto (Entrada)", "Dentro de Recinto(Entrada)", "Dentro de Recinto (Entrada)", "Dentro de Recinto Entrada"])
c_rec_out = find_col(df_asist, ["Dentro del Recinto (Salida)", "Dentro de Recinto(Salida)", "Dentro de Recinto (Salida)", "Dentro de Recinto Salida"])

if c_rec_in:
    df_asist[c_rec_in] = df_asist[c_rec_in].apply(norm_recinto)
else:
    df_asist["Dentro del Recinto (Entrada)"] = "Sin Marca"
    c_rec_in = "Dentro del Recinto (Entrada)"

if c_rec_out:
    df_asist[c_rec_out] = df_asist[c_rec_out].apply(norm_recinto)
else:
    df_asist["Dentro del Recinto (Salida)"] = "Sin Marca"
    c_rec_out = "Dentro del Recinto (Salida)"

# Fechas/hora
if "Fecha Entrada" not in df_asist.columns:
    st.error("Asistencia: falta 'Fecha Entrada' (según la detección de encabezado).")
    st.stop()

df_asist["Fecha_base"] = df_asist["Fecha Entrada"].apply(try_parse_date_any).dt.date
df_asist["HoraEntrada_str"] = df_asist["Hora Entrada"].apply(safe_time_str) if "Hora Entrada" in df_asist.columns else ""
df_asist["HoraSalida_str"] = df_asist["Hora Salida"].apply(safe_time_str) if "Hora Salida" in df_asist.columns else ""

# Inasistencias: RUT + Día
c_rut_i = find_col(df_inas, ["RUT"])
c_dia_i = find_col(df_inas, ["Día", "Dia"])
if not c_rut_i or not c_dia_i:
    st.error("Inasistencias: no encontré columnas RUT y/o Día.")
    st.stop()

df_inas["RUT_norm"] = df_inas[c_rut_i].apply(normalize_rut)
df_inas["Fecha_base"] = df_inas[c_dia_i].apply(try_parse_date_any).dt.date

# Filtro área
def maybe_filter_area(df, col="Área"):
    if only_area and col in df.columns:
        return df[df[col].astype(str).str.upper().str.contains(only_area.upper(), na=False)].copy()
    return df

df_asist = maybe_filter_area(df_asist, "Área")
df_inas = maybe_filter_area(df_inas, "Área")
df_plan = maybe_filter_area(df_plan, "Área")

# =========================
# Selector fechas (desde/hasta) basado en plan + asist + inas
# =========================
all_dates = []
if "Fecha_base" in df_asist.columns:
    all_dates += [d for d in df_asist["Fecha_base"].dropna().tolist()]
if "Fecha_base" in df_inas.columns:
    all_dates += [d for d in df_inas["Fecha_base"].dropna().tolist()]

# Planificación: columnas fechas desde la E (según doc), pero detectemos por parseo
fixed_plan_cols = ["Nombre del Colaborador", "RUT", "Área", "Supervisor"]
plan_date_cols = [c for c in df_plan.columns if c not in fixed_plan_cols]
def try_parse_plan_col_to_date(c):
    # suelen venir tipo "01/11/2025" o "2025-11-01"
    dt = pd.to_datetime(str(c), errors="coerce", dayfirst=True)
    return dt.date() if not pd.isna(dt) else None

plan_dates = [try_parse_plan_col_to_date(c) for c in plan_date_cols]
plan_dates = [d for d in plan_dates if d is not None]
all_dates += plan_dates

if not all_dates:
    st.error("No pude inferir fechas en los inputs para armar el selector de periodo.")
    st.stop()

min_d, max_d = min(all_dates), max(all_dates)

colA, colB = st.columns(2)
with colA:
    date_from = st.date_input("Desde", value=min_d, min_value=min_d, max_value=max_d)
with colB:
    date_to = st.date_input("Hasta", value=max_d, min_value=min_d, max_value=max_d)

if date_from > date_to:
    st.error("Rango inválido: 'Desde' no puede ser mayor que 'Hasta'.")
    st.stop()

# Aplicar filtro fechas
df_asist = df_asist[(df_asist["Fecha_base"] >= date_from) & (df_asist["Fecha_base"] <= date_to)].copy()
df_inas = df_inas[(df_inas["Fecha_base"] >= date_from) & (df_inas["Fecha_base"] <= date_to)].copy()

# =========================
# Planificación wide -> long, excluyendo L y vacíos
# =========================
df_plan = df_plan.copy()
df_plan["RUT_norm"] = df_plan["RUT"].apply(normalize_rut) if "RUT" in df_plan.columns else ""

date_cols = [c for c in df_plan.columns if c not in fixed_plan_cols]
plan_long = df_plan.melt(
    id_vars=[c for c in fixed_plan_cols if c in df_plan.columns],
    value_vars=date_cols,
    var_name="Fecha_col",
    value_name="Turno_Cod"
)

plan_long["Fecha"] = plan_long["Fecha_col"].apply(lambda x: try_parse_plan_col_to_date(x))
plan_long = plan_long.dropna(subset=["Fecha"])
plan_long["Fecha"] = plan_long["Fecha"].astype(object)

plan_long["RUT_norm"] = plan_long["RUT"].apply(normalize_rut) if "RUT" in plan_long.columns else ""
plan_long["Turno_Cod"] = plan_long["Turno_Cod"].astype(str).str.strip()
plan_long.loc[plan_long["Turno_Cod"].isin(["", "nan", "None"]), "Turno_Cod"] = ""

# rango seleccionado + excluir libres L
plan_long = plan_long[(plan_long["Fecha"] >= date_from) & (plan_long["Fecha"] <= date_to)].copy()
plan_long = plan_long[~plan_long["Turno_Cod"].isin(["L"])].copy()
plan_long = plan_long[plan_long["Turno_Cod"] != ""].copy()

# =========================
# Codificación: map turno -> horario -> horas planificadas
# =========================
# según doc: Sigla, Horario, Tipo, Jornada
# intentamos detectar columnas por nombre común
c_sigla = find_col(df_cod, ["Sigla", "SIGLA"])
c_hor = find_col(df_cod, ["Horario", "HORARIO"])
if not c_sigla or not c_hor:
    st.error("Codificación: no encontré columnas Sigla y Horario.")
    st.stop()

cod_map = df_cod[[c_sigla, c_hor]].copy()
cod_map.columns = ["Turno_Cod", "Horario"]
cod_map["Turno_Cod"] = cod_map["Turno_Cod"].astype(str).str.strip()
cod_map["Horas_Plan"] = cod_map["Horario"].apply(parse_horario_to_hours)

plan_long = plan_long.merge(cod_map, on="Turno_Cod", how="left")

# =========================
# Construcción Detalle (incidencias + inasistencias)
# =========================
# Asistencia: horas trabajadas aproximadas por diferencia timestamp (si hay fecha/hora)
def combine_dt(fecha_col, hora_col):
    if pd.isna(fecha_col):
        return pd.NaT
    try:
        d = pd.to_datetime(fecha_col, errors="coerce", dayfirst=True)
        if pd.isna(d):
            return pd.NaT
        if not hora_col:
            return d
        t = pd.to_datetime(hora_col, errors="coerce").time()
        return datetime(d.year, d.month, d.day, t.hour, t.minute, getattr(t, "second", 0))
    except Exception:
        return pd.NaT

# dt entrada/salida
df_asist["dt_in"] = [
    combine_dt(f, h) for f, h in zip(df_asist.get("Fecha Entrada"), df_asist.get("Hora Entrada", [""] * len(df_asist)))
]
df_asist["dt_out"] = [
    combine_dt(f, h) for f, h in zip(df_asist.get("Fecha Salida", df_asist.get("Fecha Entrada")), df_asist.get("Hora Salida", [""] * len(df_asist)))
]

# horas trabajadas
def worked_hours(dt_in, dt_out):
    if pd.isna(dt_in) or pd.isna(dt_out):
        return np.nan
    delta = (dt_out - dt_in).total_seconds() / 3600.0
    if delta < 0:
        delta += 24.0
    return delta

df_asist["Horas_Trab"] = [worked_hours(a, b) for a, b in zip(df_asist["dt_in"], df_asist["dt_out"])]

# unir con plan_long para obtener Turno_Cod / Horas_Plan
asist_merge = df_asist.merge(
    plan_long[["RUT_norm", "Fecha", "Turno_Cod", "Horario", "Horas_Plan"]],
    left_on=["RUT_norm", "Fecha_base"],
    right_on=["RUT_norm", "Fecha"],
    how="left"
)

# marcas fuera: si entrada o salida = "No"
asist_merge["Marcas_Fuera"] = (
    (asist_merge[c_rec_in].astype(str).str.upper() == "NO") |
    (asist_merge[c_rec_out].astype(str).str.upper() == "NO")
).astype(int)

# diferencia plan vs trabajada
asist_merge["Diff_h"] = (asist_merge["Horas_Plan"] - asist_merge["Horas_Trab"]).astype(float)

# regla MVP: considerar incidencia si:
# - hay turno planificado (Horas_Plan no NaN)
# - y (Diff_h >= umbral) o (Marcas_Fuera=1)
mask_inc = (
    (~asist_merge["Horas_Plan"].isna()) &
    (
        (asist_merge["Diff_h"].fillna(0) >= float(umbral_diff_h)) |
        (asist_merge["Marcas_Fuera"] == 1)
    )
)

df_inc = asist_merge[mask_inc].copy()

# columnas de identidad (según doc asistencia)
# Código/RUT/Nombre/Apellidos/Especialidad/Área/Contrato/Supervisor/Turno
c_cod_a = find_col(df_asist, ["Código", "Codigo"])
c_nom = find_col(df_asist, ["Nombre"])
c_ap1 = find_col(df_asist, ["Primer Apellido", "PrimerApellido"])
c_ap2 = find_col(df_asist, ["Segundo Apellido", "SegundoApellido"])
c_esp = find_col(df_asist, ["Especialidad"])
c_sup = find_col(df_asist, ["Supervisor"])
c_turno_txt = find_col(df_asist, ["Turno"])

def col_or_blank(df, c):
    return df[c] if c and c in df.columns else ""

df_inc_detalle = pd.DataFrame({
    "Fecha": df_inc["Fecha_base"],
    "Código": col_or_blank(df_inc, c_cod_a),
    "RUT": df_inc[c_rut_a],
    "Nombre": col_or_blank(df_inc, c_nom),
    "Primer Apellido": col_or_blank(df_inc, c_ap1),
    "Segundo Apellido": col_or_blank(df_inc, c_ap2),
    "Especialidad": col_or_blank(df_inc, c_esp),
    "Supervisor": col_or_blank(df_inc, c_sup),
    "Turno Planificado (Cod)": df_inc["Turno_Cod"],
    "Horario Planificado": df_inc["Horario"],
    "Turno Marcado": df_inc["HoraEntrada_str"].astype(str) + "-" + df_inc["HoraSalida_str"].astype(str),
    "Dentro del Recinto (Entrada)": df_inc[c_rec_in],
    "Dentro del Recinto (Salida)": df_inc[c_rec_out],
    "Tipo_Incidencia": "Marcaje/Turno",
    "Detalle": (
        "HorasPlan=" + df_inc["Horas_Plan"].astype(str) +
        " | HorasTrab=" + df_inc["Horas_Trab"].astype(str) +
        " | Diff_h=" + df_inc["Diff_h"].astype(str) +
        " | MarcasFuera=" + df_inc["Marcas_Fuera"].astype(str)
    ),
    "Clasificación Manual": "Seleccionar",
    "Minutos Retraso": "",
    "Minutos Salida Anticipada": "",
    "RUT_norm": df_inc["RUT_norm"],
})

# Inasistencias -> Detalle
c_cod_i = find_col(df_inas, ["Código", "Codigo"])
c_nom_i = find_col(df_inas, ["Nombre"])
c_ap1_i = find_col(df_inas, ["Primer Apellido", "PrimerApellido"])
c_ap2_i = find_col(df_inas, ["Segundo Apellido", "SegundoApellido"])
c_esp_i = find_col(df_inas, ["Especialidad"])
c_sup_i = find_col(df_inas, ["Supervisor"])
c_turno_i = find_col(df_inas, ["Turno"])
c_mot = find_col(df_inas, ["Motivo"])
c_obs = find_col(df_inas, ["Observación Permiso", "Observacion Permiso"])

df_inas_detalle = pd.DataFrame({
    "Fecha": df_inas["Fecha_base"],
    "Código": col_or_blank(df_inas, c_cod_i),
    "RUT": df_inas[c_rut_i],
    "Nombre": col_or_blank(df_inas, c_nom_i),
    "Primer Apellido": col_or_blank(df_inas, c_ap1_i),
    "Segundo Apellido": col_or_blank(df_inas, c_ap2_i),
    "Especialidad": col_or_blank(df_inas, c_esp_i),
    "Supervisor": col_or_blank(df_inas, c_sup_i),
    "Turno Planificado (Cod)": col_or_blank(df_inas, c_turno_i),
    "Horario Planificado": "",
    "Turno Marcado": "",
    "Dentro del Recinto (Entrada)": "Sin Marca",
    "Dentro del Recinto (Salida)": "Sin Marca",
    "Tipo_Incidencia": "Inasistencia",
    "Detalle": (
        "Motivo=" + (col_or_blank(df_inas, c_mot).astype(str) if c_mot else "") +
        " | Obs=" + (col_or_blank(df_inas, c_obs).astype(str) if c_obs else "")
    ),
    "Clasificación Manual": "Seleccionar",
    "Minutos Retraso": "",
    "Minutos Salida Anticipada": "",
    "RUT_norm": df_inas["RUT_norm"],
})

# regla extra: si motivo es P/L/V/C, sugerir clasificación
if c_mot and c_mot in df_inas.columns:
    mot = df_inas[c_mot].astype(str).str.strip().str.upper()
    df_inas_detalle.loc[mot == "P", "Clasificación Manual"] = "Permiso"
    df_inas_detalle.loc[mot == "L", "Clasificación Manual"] = "Licencia"
    df_inas_detalle.loc[mot == "V", "Clasificación Manual"] = "Vacaciones"
    df_inas_detalle.loc[mot == "C", "Clasificación Manual"] = "Compensado"

detalle = pd.concat([df_inc_detalle, df_inas_detalle], ignore_index=True)
detalle["Fecha"] = pd.to_datetime(detalle["Fecha"], errors="coerce").dt.date

# orden razonable
detalle = detalle.sort_values(["Fecha", "RUT"]).reset_index(drop=True)

# =========================
# UI: tabla Detalle (en app)
# =========================
st.subheader("Detalle (fuente editable para clasificación)")
st.caption("Este Detalle alimenta los reportes del Excel (Cumplimiento / KPIs / Ausentismo) mediante fórmulas.")

edited = st.data_editor(
    detalle.drop(columns=["RUT_norm"], errors="ignore"),
    use_container_width=True,
    num_rows="dynamic",
    column_config={
        "Tipo_Incidencia": st.column_config.SelectboxColumn(options=["Inasistencia", "Marcaje/Turno", "No Procede"]),
        "Clasificación Manual": st.column_config.SelectboxColumn(options=["Seleccionar","Injustificada","Permiso","Licencia","Vacaciones","Compensado","No Procede"]),
    }
)

# =========================
# Excel: construir workbook con fórmulas
# =========================
def write_df(ws, df: pd.DataFrame):
    ws.append(list(df.columns))
    for r in df.itertuples(index=False):
        ws.append(list(r))

def add_dropdown(ws, col_letter: str, start_row: int, end_row: int, options: list, prompt: str):
    # lista literal (ok para pocos valores)
    formula = '"' + ",".join(options) + '"'
    dv = DataValidation(type="list", formula1=formula, allow_blank=True, showDropDown=True)
    dv.promptTitle = "Seleccionar"
    dv.prompt = prompt
    ws.add_data_validation(dv)
    dv.add(f"{col_letter}{start_row}:{col_letter}{end_row}")

def build_excel(edited_df: pd.DataFrame, plan_long_df: pd.DataFrame) -> BytesIO:
    wb = Workbook()
    wb.remove(wb.active)

    # 1) Detalle
    ws_det = wb.create_sheet("Detalle")
    det_df = edited_df.copy()

    # asegurar Fecha como fecha
    det_df["Fecha"] = pd.to_datetime(det_df["Fecha"], errors="coerce").dt.date

    write_df(ws_det, det_df)
    style_header_row(ws_det, 1, CABIFY_HEADER)
    autosize_columns(ws_det)

    # Dropdowns en Detalle
    # localizar columnas
    cols = list(det_df.columns)
    def col_letter_of(col_name):
        idx = cols.index(col_name) + 1
        return get_column_letter(idx)

    n_rows = len(det_df) + 1
    if "Tipo_Incidencia" in cols:
        add_dropdown(ws_det, col_letter_of("Tipo_Incidencia"), 2, n_rows, ["Inasistencia","Marcaje/Turno","No Procede"], "Define si es Inasistencia, Marcaje/Turno o No Procede.")
    if "Clasificación Manual" in cols:
        add_dropdown(ws_det, col_letter_of("Clasificación Manual"), 2, n_rows, ["Seleccionar","Injustificada","Permiso","Licencia","Vacaciones","Compensado","No Procede"], "Clasifica el evento.")

    # 2) Planificacion_long (incluye fórmulas auxiliares)
    ws_plan = wb.create_sheet("Planificacion_long")
    pl = plan_long_df.copy()
    pl = pl[["Fecha","RUT","Nombre del Colaborador","Área","Supervisor","Turno_Cod","Horario","Horas_Plan"]].copy()
    pl["Fecha"] = pd.to_datetime(pl["Fecha"], errors="coerce").dt.date

    # helper: Ausente_Injustificada (formula por fila), Incidencia_Injustificada (formula por fila)
    # (se calcula mirando Detalle por RUT+Fecha)
    pl["Ausente_Injustificada"] = ""
    pl["Incidencia_Injustificada"] = ""

    write_df(ws_plan, pl)
    style_header_row(ws_plan, 1, CABIFY_HEADER)
    autosize_columns(ws_plan)

    # ubicar columnas en Planificacion_long para meter fórmulas por fila
    pl_cols = list(pl.columns)
    c_fecha = get_column_letter(pl_cols.index("Fecha")+1)
    c_rut = get_column_letter(pl_cols.index("RUT")+1)
    c_aus = get_column_letter(pl_cols.index("Ausente_Injustificada")+1)
    c_inc = get_column_letter(pl_cols.index("Incidencia_Injustificada")+1)

    # Detalle columnas (por letra)
    det_cols = list(det_df.columns)
    det_fecha = get_column_letter(det_cols.index("Fecha")+1)
    det_rut = get_column_letter(det_cols.index("RUT")+1)
    det_tipo = get_column_letter(det_cols.index("Tipo_Incidencia")+1)
    det_clas = get_column_letter(det_cols.index("Clasificación Manual")+1)

    # fórmula fila a fila (evitamos tablas estructuradas para mejor compatibilidad con Sheets)
    for r in range(2, len(pl)+2):
        # Ausente injustificada: existe fila en Detalle con (Fecha==A, RUT==B, Tipo==Inasistencia, Clas==Injustificada)
        ws_plan[f"{c_aus}{r}"].value = (
            f'=IF(COUNTIFS(Detalle!${det_fecha}:${det_fecha},{c_fecha}{r},'
            f'Detalle!${det_rut}:${det_rut},{c_rut}{r},'
            f'Detalle!${det_tipo}:${det_tipo},"Inasistencia",'
            f'Detalle!${det_clas}:${det_clas},"Injustificada")>0,1,0)'
        )
        # Incidencia injustificada (Marcaje/Turno)
        ws_plan[f"{c_inc}{r}"].value = (
            f'=IF(COUNTIFS(Detalle!${det_fecha}:${det_fecha},{c_fecha}{r},'
            f'Detalle!${det_rut}:${det_rut},{c_rut}{r},'
            f'Detalle!${det_tipo}:${det_tipo},"Marcaje/Turno",'
            f'Detalle!${det_clas}:${det_clas},"Injustificada")>0,1,0)'
        )

    # 3) Cumplimiento (por colaborador)
    ws_c = wb.create_sheet("Cumplimiento")
    # lista única de colaboradores desde Planificacion_long (asegura incluirlos aunque no tengan incidencias)
    base = pl[["RUT","Nombre del Colaborador","Supervisor","Área"]].drop_duplicates().copy()
    base = base.sort_values(["Nombre del Colaborador","RUT"]).reset_index(drop=True)

    out = base.copy()
    out["Turnos_Planificados"] = ""
    out["Inasistencias_Injustificadas"] = ""
    out["Incidencias_Injustificadas"] = ""
    out["Cumplimiento_%"] = ""

    write_df(ws_c, out)
    style_header_row(ws_c, 1, CABIFY_HEADER)
    autosize_columns(ws_c)

    # letras en Cumplimiento
    c_cols = list(out.columns)
    L_rut = get_column_letter(c_cols.index("RUT")+1)
    L_tp = get_column_letter(c_cols.index("Turnos_Planificados")+1)
    L_ina = get_column_letter(c_cols.index("Inasistencias_Injustificadas")+1)
    L_inci = get_column_letter(c_cols.index("Incidencias_Injustificadas")+1)
    L_cump = get_column_letter(c_cols.index("Cumplimiento_%")+1)

    # columnas en Planificacion_long para SUM/COUNT
    pl_rut = get_column_letter(pl_cols.index("RUT")+1)
    pl_fecha = get_column_letter(pl_cols.index("Fecha")+1)
    pl_ausflag = get_column_letter(pl_cols.index("Ausente_Injustificada")+1)
    pl_incflag = get_column_letter(pl_cols.index("Incidencia_Injustificada")+1)

    # Turnos planificados = COUNTIF Planificacion_long!RUT = rut
    for r in range(2, len(out)+2):
        ws_c[f"{L_tp}{r}"].value = f'=COUNTIF(Planificacion_long!${pl_rut}:${pl_rut},{L_rut}{r})'
        # inasistencias injustificadas = SUMIF Planificacion_long (flag aus) por RUT
        ws_c[f"{L_ina}{r}"].value = (
            f'=SUMIF(Planificacion_long!${pl_rut}:${pl_rut},{L_rut}{r},Planificacion_long!${pl_ausflag}:${pl_ausflag})'
        )
        ws_c[f"{L_inci}{r}"].value = (
            f'=SUMIF(Planificacion_long!${pl_rut}:${pl_rut},{L_rut}{r},Planificacion_long!${pl_incflag}:${pl_incflag})'
        )
        ws_c[f"{L_cump}{r}"].value = (
            f'=IF({L_tp}{r}=0,"",MAX(0,1-(({L_ina}{r}+{L_inci}{r})/{L_tp}{r})))'
        )

    # formato porcentaje
    for r in range(2, len(out)+2):
        ws_c[f"{L_cump}{r}"].number_format = "0.00%"

    # 4) KPIs diarios (matriz KPI x fecha)
    ws_k = wb.create_sheet("KPIs_diarios")

    fechas = pd.date_range(date_from, date_to, freq="D").date.tolist()
    kpis = [
        "Turnos_planificados",
        "Ausencias_injustificadas",
        "Incidencias_injustificadas",
        "Cumplimiento_%",
    ]

    # encabezado
    ws_k.append(["KPI"] + fechas)
    for k in kpis:
        ws_k.append([k] + [""]*len(fechas))

    style_header_row(ws_k, 1, CABIFY_HEADER)
    autosize_columns(ws_k)

    # ubicaciones (en KPIs_diarios)
    # columnas de fechas empiezan en B
    # Turnos_planificados por día: COUNTIF Planificacion_long!Fecha = fecha
    # Ausencias_injustificadas: SUMIF Planificacion_long!Fecha = fecha, sum flag aus
    # Incidencias_injustificadas: SUMIF Planificacion_long!Fecha = fecha, sum flag inc
    # Cumplimiento_%: 1 - (aus+inc)/turnos
    for j, _ in enumerate(fechas, start=2):  # B=2
        colL = get_column_letter(j)
        header_cell = f"{colL}1"  # contiene la fecha como valor

        # filas KPI: 2..(len(kpis)+1)
        # row2: Turnos
        ws_k[f"{colL}2"].value = f'=COUNTIF(Planificacion_long!${pl_fecha}:${pl_fecha},{header_cell})'
        # row3: Ausencias injust
        ws_k[f"{colL}3"].value = f'=SUMIF(Planificacion_long!${pl_fecha}:${pl_fecha},{header_cell},Planificacion_long!${pl_ausflag}:${pl_ausflag})'
        # row4: Incidencias injust
        ws_k[f"{colL}4"].value = f'=SUMIF(Planificacion_long!${pl_fecha}:${pl_fecha},{header_cell},Planificacion_long!${pl_incflag}:${pl_incflag})'
        # row5: Cumplimiento
        ws_k[f"{colL}5"].value = f'=IF({colL}2=0,"",MAX(0,1-(({colL}3+{colL}4)/{colL}2)))'
        ws_k[f"{colL}5"].number_format = "0.00%"

    # 5) Ausentismo (diario + general)
    ws_a = wb.create_sheet("Ausentismo")
    ws_a.append(["Fecha", "Horas_programadas", "Horas_ausentes_injustificadas", "Ausentismo_%"])
    style_header_row(ws_a, 1, CABIFY_HEADER)

    # Necesitamos Horas_Plan col en Planificacion_long
    pl_horas = get_column_letter(pl_cols.index("Horas_Plan")+1)

    for i, d in enumerate(fechas, start=2):
        ws_a[f"A{i}"].value = d
        # Horas programadas: SUMIF por fecha, sum Horas_Plan
        ws_a[f"B{i}"].value = f'=SUMIF(Planificacion_long!${pl_fecha}:${pl_fecha},A{i},Planificacion_long!${pl_horas}:${pl_horas})'
        # Horas ausentes injustificadas: SUMIFS(Horas_Plan, Fecha, d, Ausente_Injustificada, 1)
        ws_a[f"C{i}"].value = (
            f'=SUMIFS(Planificacion_long!${pl_horas}:${pl_horas},'
            f'Planificacion_long!${pl_fecha}:${pl_fecha},A{i},'
            f'Planificacion_long!${pl_ausflag}:${pl_ausflag},1)'
        )
        ws_a[f"D{i}"].value = f'=IF(B{i}=0,"",C{i}/B{i})'
        ws_a[f"D{i}"].number_format = "0.00%"

    # totales
    last = len(fechas) + 1
    ws_a[f"F1"].value = "Resumen"
    ws_a[f"F1"].font = Font(bold=True, color=CABIFY_ACCENT)
    ws_a[f"F2"].value = "Horas_programadas_total"
    ws_a[f"G2"].value = f"=SUM(B2:B{last})"
    ws_a[f"F3"].value = "Horas_ausentes_injust_total"
    ws_a[f"G3"].value = f"=SUM(C2:C{last})"
    ws_a[f"F4"].value = "Ausentismo_%_total"
    ws_a[f"G4"].value = f'=IF(G2=0,"",G3/G2)'
    ws_a[f"G4"].number_format = "0.00%"

    autosize_columns(ws_a)

    # Congelar paneles
    ws_det.freeze_panes = "A2"
    ws_plan.freeze_panes = "A2"
    ws_c.freeze_panes = "A2"
    ws_k.freeze_panes = "B2"
    ws_a.freeze_panes = "A2"

    # salida bytes
    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio

excel_bytes = build_excel(edited, plan_long)

st.subheader("Descarga")
st.download_button(
    "Descargar Excel (Detalle + Cumplimiento + KPIs + Ausentismo)",
    data=excel_bytes,
    file_name=f"reporte_ausentismo_incidencias_{date_from}_{date_to}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
