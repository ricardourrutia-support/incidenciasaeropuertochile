import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime, date, timedelta

from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation


# =========================
# UI
# =========================
st.set_page_config(page_title="Ausentismo e Incidencias Operativas", layout="wide")
st.title("Plataforma de Gestión de Ausentismo e Incidencias Operativas")


# =========================
# Config / Listas
# =========================
TIPO_OPTS = ["Inasistencia", "Marcaje/Turno", "No Procede"]
CLASIF_OPTS = ["Seleccionar", "Injustificada", "Permiso", "Licencia", "Vacaciones", "Compensado", "No Procede"]

CABIFY = {
    "m2": "362065",
    "m11": "FAF8FE",
    "white": "FFFFFF",
    "grid": "D9D9D9",
    "accent": "E83C96",
}

THIN = Side(style="thin", color=CABIFY["grid"])
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


# =========================
# Helpers
# =========================
def normalize_rut(x) -> str:
    if pd.isna(x):
        return ""
    return str(x).strip().upper().replace(".", "").replace(" ", "")

def try_parse_date_any(x):
    if pd.isna(x):
        return pd.NaT
    return pd.to_datetime(x, errors="coerce", dayfirst=True)

def safe_str(x):
    return "" if pd.isna(x) else str(x).strip()

def read_excel_flexible(file, header=0, sheet_name=0, skiprows=None):
    """
    Lee XLSX con openpyxl o XLS con xlrd (si está instalado).
    """
    try:
        return pd.read_excel(file, sheet_name=sheet_name, header=header, skiprows=skiprows, engine=None)
    except Exception as e:
        # Segundo intento explícito (por si Streamlit Cloud no autodetecta)
        name = getattr(file, "name", "").lower()
        if name.endswith(".xls"):
            try:
                return pd.read_excel(file, sheet_name=sheet_name, header=header, skiprows=skiprows, engine="xlrd")
            except Exception as e2:
                raise RuntimeError(
                    "No pude leer el archivo .xls. Probablemente falta xlrd.\n"
                    "Agrega 'xlrd' en requirements.txt (y vuelve a desplegar)."
                ) from e2
        else:
            try:
                return pd.read_excel(file, sheet_name=sheet_name, header=header, skiprows=skiprows, engine="openpyxl")
            except Exception as e2:
                raise RuntimeError(
                    "No pude leer el Excel. Verifica que sea .xlsx o instala el motor necesario."
                ) from e2

def read_csv_flexible(file):
    return pd.read_csv(file)

def find_col(df, candidates):
    cols = {str(c).strip().lower(): c for c in df.columns}
    for cand in candidates:
        key = str(cand).strip().lower()
        if key in cols:
            return cols[key]
    return None

def num_col(df, name):
    if name not in df.columns:
        return pd.Series([0.0] * len(df), index=df.index)
    return pd.to_numeric(df[name], errors="coerce").fillna(0.0)

def is_yes(v):
    s = str(v).strip().lower()
    return s in {"si", "sí", "s", "true", "1", "y", "yes"}

def parse_time_only(x):
    if pd.isna(x) or str(x).strip() == "":
        return None
    # soporta "19:41:00" o "4:43:34" etc.
    t = pd.to_datetime(str(x), errors="coerce")
    if pd.isna(t):
        return None
    return t.time()

def parse_range_to_times(rng: str):
    """
    Convierte "7:00:00 - 15:00:00" o "07:00-19:00" a (start_time, end_time).
    """
    if not rng or pd.isna(rng):
        return (None, None)
    s = str(rng).replace(" ", "")
    if "-" not in s:
        return (None, None)
    a, b = s.split("-", 1)

    ta = pd.to_datetime(a, errors="coerce")
    tb = pd.to_datetime(b, errors="coerce")
    if pd.isna(ta) or pd.isna(tb):
        return (None, None)
    return (ta.time(), tb.time())

def combine_date_time(d: date, t):
    if d is None or pd.isna(d) or t is None:
        return None
    if isinstance(d, pd.Timestamp):
        d = d.date()
    return datetime(d.year, d.month, d.day, t.hour, t.minute, t.second)

def hours_between(dt_start: datetime, dt_end: datetime) -> float:
    if dt_start is None or dt_end is None:
        return 0.0
    delta = dt_end - dt_start
    return delta.total_seconds() / 3600.0

def ensure_overnight(start_dt, end_dt):
    if start_dt is None or end_dt is None:
        return (start_dt, end_dt)
    if end_dt < start_dt:
        end_dt = end_dt + timedelta(days=1)
    return (start_dt, end_dt)

def style_sheet_table(ws, header_row=1):
    header_fill = PatternFill("solid", fgColor=CABIFY["m2"])
    header_font = Font(color=CABIFY["white"], bold=True)
    for cell in ws[header_row]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER

    for row in ws.iter_rows(min_row=header_row + 1, max_row=ws.max_row, max_col=ws.max_column):
        for cell in row:
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    ws.freeze_panes = "A2"

    # autosize simple
    for col_cells in ws.columns:
        max_len = 10
        col_letter = col_cells[0].column_letter
        for c in col_cells[:300]:
            v = "" if c.value is None else str(c.value)
            max_len = max(max_len, len(v))
        ws.column_dimensions[col_letter].width = min(max(10, max_len + 2), 45)


# =========================
# Sidebar Inputs (NUEVOS)
# =========================
with st.sidebar:
    st.header("Inputs (nuevo esquema)")

    f_asistencia = st.file_uploader("1) Reporte de Asistencia (XLS/XLSX)", type=["xls", "xlsx"])
    f_inasist = st.file_uploader("2) Reporte de Inasistencias (XLS/XLSX)", type=["xls", "xlsx"])
    f_planif = st.file_uploader("3) Planificación de Turnos (CSV)", type=["csv"])
    f_codif = st.file_uploader("4) Codificación de Turnos (CSV)", type=["csv"])

    st.divider()
    st.subheader("Filtros / Reglas")
    only_area = st.text_input("Filtrar Área (opcional)", value="AEROPUERTO")
    min_diff_h = st.number_input(
        "Diferencia horaria mínima (horas) para considerar incidencia",
        value=0.5, step=0.25, min_value=0.0
    )
    st.caption("Se filtran incidencias con |Horas planificadas - Horas trabajadas| >= umbral.")
    st.divider()
    st.subheader("Periodo")
    st.caption("Rango se aplica sobre planificación, y se cruza con asistencia/inasistencia.")

if not all([f_asistencia, f_inasist, f_planif, f_codif]):
    st.info("Sube los 4 archivos para comenzar.")
    st.stop()


# =========================
# Load files according to spec in doc
# - Asistencia: encabezados desde A3 (header=2)
# - Inasistencias: header en fila 1 (header=0)
# - Planificación: CSV normal
# - Codificación: CSV normal
# =========================
try:
    df_asist = read_excel_flexible(f_asistencia, header=2)  # A3
    df_inas = read_excel_flexible(f_inasist, header=0)      # A1
    df_plan = read_csv_flexible(f_planif)
    df_cod = read_csv_flexible(f_codif)
except Exception as e:
    st.error(str(e))
    st.stop()

# Normaliza columnas (strip)
df_asist.columns = [str(c).strip() for c in df_asist.columns]
df_inas.columns = [str(c).strip() for c in df_inas.columns]
df_plan.columns = [str(c).strip() for c in df_plan.columns]
df_cod.columns = [str(c).strip() for c in df_cod.columns]


# =========================
# Map codificación: Sigla -> Horario
# =========================
c_sigla = find_col(df_cod, ["Sigla"])
c_hor = find_col(df_cod, ["Horario"])
if not c_sigla or not c_hor:
    st.error("No encontré columnas Sigla/Horario en Codificación.")
    st.stop()

cod_map = (
    df_cod[[c_sigla, c_hor]]
    .dropna()
    .assign(Sigla=lambda x: x[c_sigla].astype(str).str.strip(),
            Horario=lambda x: x[c_hor].astype(str).str.strip())
)
turno_to_hor = dict(zip(cod_map["Sigla"], cod_map["Horario"]))


# =========================
# Planificación (CSV ancho) -> largo
# Encabezados: Nombre del Colaborador, RUT, Área, Supervisor, luego fechas E..n
# "L" = Libre, vacío = sin turno activo
# =========================
col_name = find_col(df_plan, ["Nombre del Colaborador"])
col_rut = find_col(df_plan, ["RUT"])
col_area = find_col(df_plan, ["Área"])
col_sup = find_col(df_plan, ["Supervisor"])

for need, label in [(col_name, "Nombre del Colaborador"), (col_rut, "RUT"), (col_area, "Área"), (col_sup, "Supervisor")]:
    if not need:
        st.error(f"Falta columna '{label}' en Planificación.")
        st.stop()

df_plan["RUT_norm"] = df_plan[col_rut].apply(normalize_rut)

fixed = [col_name, col_rut, col_area, col_sup, "RUT_norm"]
date_cols = [c for c in df_plan.columns if c not in fixed]

df_pl_long = df_plan.melt(
    id_vars=fixed,
    value_vars=date_cols,
    var_name="Fecha",
    value_name="Turno_Cod"
)
df_pl_long["Fecha_dt"] = df_pl_long["Fecha"].apply(try_parse_date_any)
df_pl_long["Turno_Cod"] = df_pl_long["Turno_Cod"].astype(str).str.strip()
df_pl_long.loc[df_pl_long["Turno_Cod"].isin(["", "nan", "NaT", "None"]), "Turno_Cod"] = ""

# filtro área
if only_area:
    df_pl_long = df_pl_long[df_pl_long[col_area].astype(str).str.upper().str.contains(only_area.upper(), na=False)].copy()

# selector fechas basado en planificación
min_dt = df_pl_long["Fecha_dt"].min()
max_dt = df_pl_long["Fecha_dt"].max()
if pd.isna(min_dt) or pd.isna(max_dt):
    st.error("No pude parsear las fechas en encabezados del CSV de planificación.")
    st.stop()

c1, c2 = st.columns(2)
with c1:
    start_date = st.date_input("Desde", value=min_dt.date())
with c2:
    end_date = st.date_input("Hasta", value=max_dt.date())

if start_date > end_date:
    st.error("Rango de fechas inválido (Desde > Hasta).")
    st.stop()

start_dt = pd.to_datetime(start_date)
end_dt = pd.to_datetime(end_date)

df_pl_long = df_pl_long[(df_pl_long["Fecha_dt"] >= start_dt) & (df_pl_long["Fecha_dt"] <= end_dt)].copy()

# turnos activos: no vacío y no libre (L)
df_pl_long["Es_Libre"] = df_pl_long["Turno_Cod"].astype(str).str.upper().eq("L")
df_pl_act = df_pl_long[(df_pl_long["Turno_Cod"] != "") & (~df_pl_long["Es_Libre"])].copy()

# horario planificado desde codificación
df_pl_act["Horario_Plan"] = df_pl_act["Turno_Cod"].map(turno_to_hor).fillna("")
df_pl_act[["PlanStart_t", "PlanEnd_t"]] = df_pl_act["Horario_Plan"].apply(lambda x: pd.Series(parse_range_to_times(x)))


# =========================
# Base válidos: SOLO RUTs en planificación (evita “fantasmas”)
# =========================
valid_ruts = set(df_pl_long["RUT_norm"].dropna().unique().tolist())


# =========================
# Asistencia: columnas según doc (A3..)
# Código, RUT, Nombre, Primer Apellido, Segundo Apellido, Especialidad, Área, Supervisor,
# Turno (rango), Fecha Entrada, Hora Entrada, Fecha Salida, Hora Salida,
# Dentro del Recinto (Entrada/Salida)
# =========================
# normaliza rut
c_rut_a = find_col(df_asist, ["RUT"])
if not c_rut_a:
    st.error("No encontré columna RUT en Asistencia.")
    st.stop()
df_asist["RUT_norm"] = df_asist[c_rut_a].apply(normalize_rut)

# filtro por área y ruts válidos
if only_area:
    c_area_a = find_col(df_asist, ["Área"])
    if c_area_a:
        df_asist = df_asist[df_asist[c_area_a].astype(str).str.upper().str.contains(only_area.upper(), na=False)].copy()

df_asist = df_asist[df_asist["RUT_norm"].isin(valid_ruts)].copy()

# fecha base: Fecha Entrada
c_fent = find_col(df_asist, ["Fecha Entrada"])
c_hent = find_col(df_asist, ["Hora Entrada"])
c_fsal = find_col(df_asist, ["Fecha Salida"])
c_hsal = find_col(df_asist, ["Hora Salida"])

if not c_fent:
    st.error("No encontré columna 'Fecha Entrada' en Asistencia.")
    st.stop()

df_asist["Fecha_base"] = df_asist[c_fent].apply(try_parse_date_any)
df_asist = df_asist[(df_asist["Fecha_base"] >= start_dt) & (df_asist["Fecha_base"] <= end_dt)].copy()

df_asist["Clave_RUT_Fecha"] = df_asist["RUT_norm"].astype(str) + "_" + df_asist["Fecha_base"].dt.strftime("%Y-%m-%d").fillna("")


# =========================
# Inasistencias: según doc (A1 encabezado, A2 datos)
# Campos: RUT, Nombre, Apellidos, Especialidad, Área, Turno, Supervisor, Día, Motivo
# =========================
c_rut_i = find_col(df_inas, ["RUT"])
c_dia = find_col(df_inas, ["Día"])
if not c_rut_i or not c_dia:
    st.error("No encontré columnas RUT/Día en Inasistencias.")
    st.stop()

df_inas["RUT_norm"] = df_inas[c_rut_i].apply(normalize_rut)
if only_area:
    c_area_i = find_col(df_inas, ["Área"])
    if c_area_i:
        df_inas = df_inas[df_inas[c_area_i].astype(str).str.upper().str.contains(only_area.upper(), na=False)].copy()

df_inas = df_inas[df_inas["RUT_norm"].isin(valid_ruts)].copy()
df_inas["Fecha_base"] = df_inas[c_dia].apply(try_parse_date_any)
df_inas = df_inas[(df_inas["Fecha_base"] >= start_dt) & (df_inas["Fecha_base"] <= end_dt)].copy()
df_inas["Clave_RUT_Fecha"] = df_inas["RUT_norm"].astype(str) + "_" + df_inas["Fecha_base"].dt.strftime("%Y-%m-%d").fillna("")


# =========================
# Turnos planificados por rut-fecha (para calcular horas planificadas y ventanas)
# =========================
df_pl_act["Clave_RUT_Fecha"] = df_pl_act["RUT_norm"].astype(str) + "_" + df_pl_act["Fecha_dt"].dt.strftime("%Y-%m-%d").fillna("")
plan_lookup = df_pl_act.set_index("Clave_RUT_Fecha")[["Turno_Cod", "Horario_Plan", "PlanStart_t", "PlanEnd_t"]].to_dict(orient="index")


def planned_window_for(key: str):
    info = plan_lookup.get(key)
    if not info:
        return (None, None, "", "")
    return (info.get("PlanStart_t"), info.get("PlanEnd_t"), info.get("Turno_Cod", ""), info.get("Horario_Plan", ""))

def planned_hours_for(key: str):
    ps, pe, _, _ = planned_window_for(key)
    if ps is None or pe is None:
        return 0.0
    # duración (overnight)
    dt0 = datetime(2000, 1, 1, ps.hour, ps.minute, ps.second)
    dt1 = datetime(2000, 1, 1, pe.hour, pe.minute, pe.second)
    if dt1 < dt0:
        dt1 += timedelta(days=1)
    return round((dt1 - dt0).total_seconds() / 3600.0, 2)


# =========================
# Construcción DETALLE (nuevo pipeline)
# - Inasistencias: todas entran
# - Marcaje/Turno: se considera incidencia si |horas_plan - horas_trab| >= umbral
# =========================
def get_text(df, candidates, default=""):
    c = find_col(df, candidates)
    if not c:
        return pd.Series([default] * len(df), index=df.index)
    return df[c].astype(str).fillna(default)

# Asistencia: calcula horas trabajadas + minutos atraso/salida anticipada (según turno planificado)
as_rows = []

name_a = get_text(df_asist, ["Nombre"], "")
pa_a = get_text(df_asist, ["Primer Apellido"], "")
sa_a = get_text(df_asist, ["Segundo Apellido"], "")
esp_a = get_text(df_asist, ["Especialidad"], "")
sup_a = get_text(df_asist, ["Supervisor"], "")
turno_a = get_text(df_asist, ["Turno"], "")  # viene como rango (ej 20:00-07:00)

rec_in = get_text(df_asist, ["Dentro del Recinto (Entrada)"], "")
rec_out = get_text(df_asist, ["Dentro del Recinto (Salida)"], "")

for idx, r in df_asist.iterrows():
    key = r["Clave_RUT_Fecha"]
    ps, pe, plan_code, plan_hor = planned_window_for(key)

    # Ventana planificada con fecha base
    d0 = r["Fecha_base"].date() if not pd.isna(r["Fecha_base"]) else None
    plan_start = combine_date_time(d0, ps) if ps else None
    plan_end = combine_date_time(d0, pe) if pe else None
    plan_start, plan_end = ensure_overnight(plan_start, plan_end)

    # Marca real
    d_in = try_parse_date_any(r.get(c_fent)).date() if c_fent else d0
    t_in = parse_time_only(r.get(c_hent)) if c_hent else None
    d_out = try_parse_date_any(r.get(c_fsal)).date() if c_fsal else d0
    t_out = parse_time_only(r.get(c_hsal)) if c_hsal else None

    real_start = combine_date_time(d_in, t_in) if d_in and t_in else None
    real_end = combine_date_time(d_out, t_out) if d_out and t_out else None
    real_start, real_end = ensure_overnight(real_start, real_end)

    horas_trab = round(hours_between(real_start, real_end), 2) if (real_start and real_end) else 0.0
    horas_plan = planned_hours_for(key)

    # Incidencia preliminar por diferencia de duración
    diff = abs(horas_plan - horas_trab) if horas_plan > 0 else 0.0
    if diff < float(min_diff_h):
        continue  # no es incidencia (según nueva regla)

    # Minutos atraso / salida anticipada (si hay ventana planificada)
    min_atraso = 0
    min_salida = 0
    if plan_start and real_start:
        min_atraso = int(max(0, (real_start - plan_start).total_seconds() // 60))
    if plan_end and real_end:
        min_salida = int(max(0, (plan_end - real_end).total_seconds() // 60))

    turno_marcado = ""
    if real_start and real_end:
        turno_marcado = f"{real_start.strftime('%H:%M')}-{real_end.strftime('%H:%M')}"

    as_rows.append({
        "Fecha": d0,
        "Nombre": safe_str(name_a.loc[idx]),
        "Primer Apellido": safe_str(pa_a.loc[idx]),
        "Segundo Apellido": safe_str(sa_a.loc[idx]),
        "RUT": safe_str(r.get(c_rut_a)),
        "Especialidad": safe_str(esp_a.loc[idx]),
        "Turno": plan_hor if plan_hor else safe_str(turno_a.loc[idx]),
        "Supervisor": safe_str(sup_a.loc[idx]),
        "Turno Marcado": turno_marcado,
        "Dentro del Recinto (Entrada)": safe_str(rec_in.loc[idx]) if safe_str(rec_in.loc[idx]) else "",
        "Dentro del Recinto (Salida)": safe_str(rec_out.loc[idx]) if safe_str(rec_out.loc[idx]) else "",
        "Detalle": f"HorasPlan={horas_plan} | HorasTrab={horas_trab} | Diff={round(diff,2)}",
        "Tipo_Incidencia": "Marcaje/Turno",
        "Clasificación Manual": "Seleccionar",
        "Minutos Retraso": min_atraso,
        "Minutos Salida Anticipada": min_salida,
        "Horas Planificadas Día": horas_plan,
        "Clave_RUT_Fecha": key,
    })

df_det_as = pd.DataFrame(as_rows)

# Inasistencias: todas
inas_rows = []
name_i = get_text(df_inas, ["Nombre"], "")
pa_i = get_text(df_inas, ["Primer Apellido"], "")
sa_i = get_text(df_inas, ["Segundo Apellido"], "")
esp_i = get_text(df_inas, ["Especialidad"], "")
sup_i = get_text(df_inas, ["Supervisor"], "")
turno_i = get_text(df_inas, ["Turno"], "")
mot_i = get_text(df_inas, ["Motivo"], "")

mot_map = {"-": "Injustificada", "P": "Permiso", "L": "Licencia", "V": "Vacaciones", "C": "Compensado"}

for idx, r in df_inas.iterrows():
    d0 = r["Fecha_base"].date() if not pd.isna(r["Fecha_base"]) else None
    key = r["Clave_RUT_Fecha"]
    horas_plan = planned_hours_for(key)
    m = safe_str(mot_i.loc[idx]).upper()
    clas = mot_map.get(m, "Seleccionar")

    ps, pe, plan_code, plan_hor = planned_window_for(key)

    inas_rows.append({
        "Fecha": d0,
        "Nombre": safe_str(name_i.loc[idx]),
        "Primer Apellido": safe_str(pa_i.loc[idx]),
        "Segundo Apellido": safe_str(sa_i.loc[idx]),
        "RUT": safe_str(r.get(c_rut_i)),
        "Especialidad": safe_str(esp_i.loc[idx]),
        "Turno": plan_hor if plan_hor else safe_str(turno_i.loc[idx]),
        "Supervisor": safe_str(sup_i.loc[idx]),
        "Turno Marcado": "Sin Marca",
        "Dentro del Recinto (Entrada)": "Sin Marca",
        "Dentro del Recinto (Salida)": "Sin Marca",
        "Detalle": f"Motivo={safe_str(mot_i.loc[idx])}",
        "Tipo_Incidencia": "Inasistencia",
        "Clasificación Manual": clas if clas else "Seleccionar",
        "Minutos Retraso": 0,
        "Minutos Salida Anticipada": 0,
        "Horas Planificadas Día": horas_plan,
        "Clave_RUT_Fecha": key,
    })

df_det_in = pd.DataFrame(inas_rows)

df_det = pd.concat([df_det_as, df_det_in], ignore_index=True)

# Orden solicitado (sin columnas internas)
cols_order = [
    "Fecha", "Nombre", "Primer Apellido", "Segundo Apellido", "RUT",
    "Especialidad", "Turno", "Supervisor",
    "Turno Marcado",
    "Dentro del Recinto (Entrada)", "Dentro del Recinto (Salida)",
    "Detalle", "Tipo_Incidencia", "Clasificación Manual",
    "Minutos Retraso", "Minutos Salida Anticipada",
    "Horas Planificadas Día",
]
df_det = df_det[cols_order].copy()
df_det = df_det.sort_values(["Fecha", "RUT"], na_position="last").reset_index(drop=True)


# =========================
# UI: Editor
# =========================
st.subheader("Reporte Total de Incidencias (para clasificar)")

edited = st.data_editor(
    df_det,
    use_container_width=True,
    num_rows="dynamic",
    column_config={
        "Fecha": st.column_config.DateColumn(format="DD/MM/YYYY"),
        "Tipo_Incidencia": st.column_config.SelectboxColumn(options=TIPO_OPTS, required=True),
        "Clasificación Manual": st.column_config.SelectboxColumn(options=CLASIF_OPTS, required=True),
        "Minutos Retraso": st.column_config.NumberColumn(min_value=0, step=1),
        "Minutos Salida Anticipada": st.column_config.NumberColumn(min_value=0, step=1),
    },
)

# =========================
# Resumen dinámico (app)
# (basado en minutos manuales)
# =========================
st.subheader("Resumen dinámico (en app)")

ed = edited.copy()
ed["Minutos Retraso"] = pd.to_numeric(ed["Minutos Retraso"], errors="coerce").fillna(0)
ed["Minutos Salida Anticipada"] = pd.to_numeric(ed["Minutos Salida Anticipada"], errors="coerce").fillna(0)
ed["Horas Planificadas Día"] = pd.to_numeric(ed["Horas Planificadas Día"], errors="coerce").fillna(0)

inas_inj = ed[(ed["Tipo_Incidencia"] == "Inasistencia") & (ed["Clasificación Manual"] == "Injustificada")]
inc_inj = ed[(ed["Tipo_Incidencia"] == "Marcaje/Turno") & (ed["Clasificación Manual"] == "Injustificada")]

resumen = pd.DataFrame([
    ["Inasistencia (Total)", int((ed["Tipo_Incidencia"] == "Inasistencia").sum())],
    ["Marcaje/Turno (Total)", int((ed["Tipo_Incidencia"] == "Marcaje/Turno").sum())],
    ["Inasistencia Injustificada (Count)", int(len(inas_inj))],
    ["Retraso Injustificado (Horas)", round(float(inc_inj["Minutos Retraso"].sum() / 60), 2)],
    ["Salida Anticipada Injustificada (Horas)", round(float(inc_inj["Minutos Salida Anticipada"].sum() / 60), 2)],
], columns=["KPI", "Valor"])

st.dataframe(resumen, use_container_width=True)


# =========================
# KPI diarios (matriz) - en app
# Turnos planificados se toman de planificación activa del periodo
# =========================
st.subheader("KPIs diarios (matriz)")

dates = pd.date_range(start_dt, end_dt, freq="D").date
planned_daily = (
    df_pl_act.assign(Fecha_d=df_pl_act["Fecha_dt"].dt.date)
    .groupby("Fecha_d").size().to_dict()
)

aus_inj_daily = (
    inas_inj.assign(Fecha_d=pd.to_datetime(inas_inj["Fecha"]).dt.date)
    .groupby("Fecha_d").size().to_dict()
)

inc_inj_daily = (
    inc_inj.assign(Fecha_d=pd.to_datetime(inc_inj["Fecha"]).dt.date)
    .groupby("Fecha_d").size().to_dict()
)

mat = pd.DataFrame(index=[
    "Turnos_planificados",
    "Ausencias_Injustificadas",
    "Incidencias_Injustificadas",
    "Cumplimiento_%"
], columns=dates)

for d in dates:
    tp = int(planned_daily.get(d, 0))
    ai = int(aus_inj_daily.get(d, 0))
    ii = int(inc_inj_daily.get(d, 0))
    mat.loc["Turnos_planificados", d] = tp
    mat.loc["Ausencias_Injustificadas", d] = ai
    mat.loc["Incidencias_Injustificadas", d] = ii
    mat.loc["Cumplimiento_%", d] = "" if tp == 0 else round((tp - ai - ii) / tp, 4)

st.dataframe(mat, use_container_width=True)


# =========================
# Export Excel (Cabify + dropdowns + fórmulas activas)
# - Detalle (editable)
# - Resumen (fórmulas)
# - KPIs_diarios (fórmulas)
# - Listas (oculto)
# =========================
def build_excel(df_detalle: pd.DataFrame, start_dt, end_dt, planned_daily: dict):
    wb = Workbook()
    wb.remove(wb.active)

    # Listas
    ws_l = wb.create_sheet("Listas")
    ws_l["A1"] = "Tipo_Incidencia"
    for i, v in enumerate(TIPO_OPTS, start=2):
        ws_l[f"A{i}"] = v
    ws_l["C1"] = "Clasificación Manual"
    for i, v in enumerate(CLASIF_OPTS, start=2):
        ws_l[f"C{i}"] = v
    ws_l.sheet_state = "hidden"

    # Detalle
    ws = wb.create_sheet("Detalle")
    df_out = df_detalle.copy()
    df_out["Fecha"] = pd.to_datetime(df_out["Fecha"], errors="coerce")
    for r in dataframe_to_rows(df_out, index=False, header=True):
        ws.append(r)
    style_sheet_table(ws)

    # fecha corta
    fecha_col = list(df_out.columns).index("Fecha") + 1
    for rr in range(2, ws.max_row + 1):
        ws.cell(rr, fecha_col).number_format = "DD/MM/YYYY"

    # Validaciones en Detalle
    col_tipo = list(df_out.columns).index("Tipo_Incidencia") + 1
    col_clas = list(df_out.columns).index("Clasificación Manual") + 1

    dv_tipo = DataValidation(type="list", formula1="=Listas!$A$2:$A$4", allow_blank=False)
    dv_clas = DataValidation(type="list", formula1="=Listas!$C$2:$C$8", allow_blank=False)
    ws.add_data_validation(dv_tipo)
    ws.add_data_validation(dv_clas)

    dv_tipo.add(f"{ws.cell(2, col_tipo).coordinate}:{ws.cell(ws.max_row, col_tipo).coordinate}")
    dv_clas.add(f"{ws.cell(2, col_clas).coordinate}:{ws.cell(ws.max_row, col_clas).coordinate}")

    # Resumen (fórmulas activas)
    ws_r = wb.create_sheet("Resumen")
    ws_r.append(["KPI", "Valor"])
    # Column letters based on Detalle structure:
    # A Fecha, M Tipo_Incidencia, N Clasificación Manual, O Minutos Retraso, P Minutos Salida, Q Horas Planificadas Día (si cambia, se ajusta por índice abajo)
    cols = list(df_out.columns)
    colA = "A"
    colTipoL = ws.cell(1, cols.index("Tipo_Incidencia") + 1).column_letter
    colClasL = ws.cell(1, cols.index("Clasificación Manual") + 1).column_letter
    colMinRetL = ws.cell(1, cols.index("Minutos Retraso") + 1).column_letter
    colMinSalL = ws.cell(1, cols.index("Minutos Salida Anticipada") + 1).column_letter
    colHorasPlanL = ws.cell(1, cols.index("Horas Planificadas Día") + 1).column_letter

    rows = [
        ("Inasistencia (Total)", f'=COUNTIF(Detalle!{colTipoL}:{colTipoL},"Inasistencia")'),
        ("Marcaje/Turno (Total)", f'=COUNTIF(Detalle!{colTipoL}:{colTipoL},"Marcaje/Turno")'),
        ("Inasistencia Injustificada (Count)", f'=COUNTIFS(Detalle!{colTipoL}:{colTipoL},"Inasistencia",Detalle!{colClasL}:{colClasL},"Injustificada")'),
        ("Retraso Injustificado (Horas)", f'=SUMIFS(Detalle!{colMinRetL}:{colMinRetL},Detalle!{colTipoL}:{colTipoL},"Marcaje/Turno",Detalle!{colClasL}:{colClasL},"Injustificada")/60'),
        ("Salida Anticipada Injustificada (Horas)", f'=SUMIFS(Detalle!{colMinSalL}:{colMinSalL},Detalle!{colTipoL}:{colTipoL},"Marcaje/Turno",Detalle!{colClasL}:{colClasL},"Injustificada")/60'),
        ("Ausentismo (aprox)", f'=IF(SUM(Detalle!{colHorasPlanL}:{colHorasPlanL})=0,"",('
                               f'SUMIFS(Detalle!{colHorasPlanL}:{colHorasPlanL},Detalle!{colTipoL}:{colTipoL},"Inasistencia",Detalle!{colClasL}:{colClasL},"<>Seleccionar")'
                               f'+(SUMIFS(Detalle!{colMinRetL}:{colMinRetL},Detalle!{colTipoL}:{colTipoL},"Marcaje/Turno",Detalle!{colClasL}:{colClasL},"Injustificada")'
                               f'+SUMIFS(Detalle!{colMinSalL}:{colMinSalL},Detalle!{colTipoL}:{colTipoL},"Marcaje/Turno",Detalle!{colClasL}:{colClasL},"Injustificada"))/60)'
                               f'/SUM(Detalle!{colHorasPlanL}:{colHorasPlanL}))')
    ]
    for k, f in rows:
        ws_r.append([k, f])
    style_sheet_table(ws_r)
    ws_r["B7"].number_format = "0.00%"

    # KPIs diarios (fórmulas + planificados fijos desde app)
    ws_k = wb.create_sheet("KPIs_diarios")
    ws_k.cell(row=1, column=1, value="KPI")
    all_dates = list(pd.date_range(start_dt, end_dt, freq="D"))
    for j, d in enumerate(all_dates, start=2):
        ws_k.cell(row=1, column=j, value=d.to_pydatetime())
        ws_k.cell(row=1, column=j).number_format = "DD/MM/YYYY"

    kpis = ["Turnos_planificados", "Ausencias_Injustificadas", "Incidencias_Injustificadas", "Cumplimiento_%"]
    for i, k in enumerate(kpis, start=2):
        ws_k.cell(row=i, column=1, value=k)

    for j, d in enumerate(all_dates, start=2):
        d_cell = ws_k.cell(row=1, column=j).coordinate
        # planificados fijos (desde planificación, no desde Detalle)
        ws_k.cell(row=2, column=j, value=int(planned_daily.get(d.date(), 0)))

        ws_k.cell(row=3, column=j, value=f'=COUNTIFS(Detalle!{colA}:{colA},{d_cell},Detalle!{colTipoL}:{colTipoL},"Inasistencia",Detalle!{colClasL}:{colClasL},"Injustificada")')
        ws_k.cell(row=4, column=j, value=f'=COUNTIFS(Detalle!{colA}:{colA},{d_cell},Detalle!{colTipoL}:{colTipoL},"Marcaje/Turno",Detalle!{colClasL}:{colClasL},"Injustificada")')
        tp_cell = ws_k.cell(row=2, column=j).coordinate
        ai_cell = ws_k.cell(row=3, column=j).coordinate
        ii_cell = ws_k.cell(row=4, column=j).coordinate
        ws_k.cell(row=5, column=j, value=f'=IF({tp_cell}=0,"",1-(({ai_cell}+{ii_cell})/{tp_cell}))')
        ws_k.cell(row=5, column=j).number_format = "0.00%"

    style_sheet_table(ws_k)
    ws_k.freeze_panes = "B2"
    ws_k.column_dimensions["A"].width = 28

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output


st.subheader("Descarga")
excel_bytes = build_excel(edited, start_dt, end_dt, planned_daily)
st.download_button(
    "Descargar Excel consolidado (dropdowns + fórmulas)",
    data=excel_bytes,
    file_name="reporte_ausentismo_incidencias.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
