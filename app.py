import streamlit as st
import pandas as pd
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
# Config
# =========================
TIPO_OPTS = ["Inasistencia", "Marcaje/Turno", "No Procede"]
CLASIF_OPTS = ["Seleccionar", "Injustificada", "Permiso", "Licencia", "Vacaciones", "Compensado", "No Procede"]

CABIFY = {"header": "362065", "white": "FFFFFF", "grid": "D9D9D9"}
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

def _norm_colname(s: str) -> str:
    s = "" if s is None else str(s)
    # limpia espacios raros
    s = s.replace("\u00A0", " ").replace("\u2007", " ").replace("\u202F", " ")
    s = s.strip().lower()
    # acentos
    s = (s.replace("á","a").replace("é","e").replace("í","i")
           .replace("ó","o").replace("ú","u").replace("ñ","n"))
    # separadores
    for ch in [" ", ".", "-", "_", "\n", "\t", "\r", ":", ";", ","]:
        s = s.replace(ch, "")
    s = s.replace("(", "").replace(")", "")
    return s

def find_col(df: pd.DataFrame, candidates: list[str]):
    colmap = {_norm_colname(c): c for c in df.columns}
    for cand in candidates:
        key = _norm_colname(cand)
        if key in colmap:
            return colmap[key]
    return None

def _clean_cell(x):
    if pd.isna(x):
        return ""
    s = str(x)
    s = s.replace("\u00A0", " ").replace("\u2007", " ").replace("\u202F", " ")
    return s.strip()

def _ffill_row(values):
    out = []
    last = ""
    for v in values:
        s = _clean_cell(v)
        if s:
            last = s
            out.append(s)
        else:
            out.append(last)
    return out


# =========================
# Excel readers
# =========================
def _read_excel_raw_noheader(file, sheet_name=0):
    name = getattr(file, "name", "").lower()
    if name.endswith(".xls"):
        return pd.read_excel(file, sheet_name=sheet_name, header=None, engine="xlrd")
    return pd.read_excel(file, sheet_name=sheet_name, header=None, engine="openpyxl")

def read_asistencia_b1b2(file, sheet_name=0):
    """
    Asistencia:
    - Encabezado en B1+B2 (merged)
    - Datos desde B3
    - Ignora columna A
    """
    raw = _read_excel_raw_noheader(file, sheet_name=sheet_name)
    if len(raw) < 3:
        raise RuntimeError("Asistencia: menos de 3 filas. No puedo construir encabezado B1/B2.")

    row1 = _ffill_row(raw.iloc[0].tolist())[1:]  # desde B
    row2 = _ffill_row(raw.iloc[1].tolist())[1:]  # desde B

    cols = []
    for a, b in zip(row1, row2):
        if a and b and _norm_colname(a) != _norm_colname(b):
            cols.append(f"{a} {b}".strip())
        elif b:
            cols.append(b)
        else:
            cols.append(a)

    cols = [(_clean_cell(c) if _clean_cell(c) else f"COL_{i+1}") for i, c in enumerate(cols)]
    df = raw.iloc[2:, 1:].copy()  # desde fila 3, col B
    df.columns = cols
    df = df.dropna(how="all")
    df.columns = [_clean_cell(c) for c in df.columns]
    return df, raw

def read_inasistencia_detect(file, sheet_name=0, max_scan_rows=80):
    raw = _read_excel_raw_noheader(file, sheet_name=sheet_name)

    def row_has(row_vals, must):
        keys = {_norm_colname(v) for v in row_vals}
        return all(_norm_colname(m) in keys for m in must)

    header_row = None
    for i in range(min(max_scan_rows, len(raw))):
        if row_has(raw.iloc[i].tolist(), ("RUT", "Día")) or row_has(raw.iloc[i].tolist(), ("RUT", "Dia")):
            header_row = i
            break

    if header_row is None:
        for i in range(min(max_scan_rows, len(raw))):
            if row_has(raw.iloc[i].tolist(), ("RUT",)):
                header_row = i
                break

    if header_row is None:
        raise RuntimeError("Inasistencias: no pude detectar encabezado (no encontré RUT / Día).")

    cols = [_clean_cell(c) for c in raw.iloc[header_row].tolist()]
    df = raw.iloc[header_row + 1:].copy()
    df.columns = cols
    df = df.dropna(how="all")
    df.columns = [_clean_cell(c) for c in df.columns]
    return df, raw


# =========================
# Turnos / horas
# =========================
def parse_time_only(x):
    if pd.isna(x) or str(x).strip() == "":
        return None
    t = pd.to_datetime(str(x), errors="coerce")
    if pd.isna(t):
        return None
    return t.time()

def parse_range_to_times(rng: str):
    if rng is None or pd.isna(rng):
        return (None, None)
    s = str(rng).strip()
    if "-" not in s:
        return (None, None)
    a, b = s.split("-", 1)
    ta = pd.to_datetime(a.strip(), errors="coerce")
    tb = pd.to_datetime(b.strip(), errors="coerce")
    if pd.isna(ta) or pd.isna(tb):
        return (None, None)
    return (ta.time(), tb.time())

def combine_date_time(d: date, t):
    if d is None or pd.isna(d) or t is None:
        return None
    if isinstance(d, pd.Timestamp):
        d = d.date()
    return datetime(d.year, d.month, d.day, t.hour, t.minute, t.second)

def ensure_overnight(start_dt, end_dt):
    if start_dt is None or end_dt is None:
        return (start_dt, end_dt)
    if end_dt < start_dt:
        end_dt = end_dt + timedelta(days=1)
    return (start_dt, end_dt)

def hours_between(dt_start: datetime, dt_end: datetime) -> float:
    if dt_start is None or dt_end is None:
        return 0.0
    return (dt_end - dt_start).total_seconds() / 3600.0


# =========================
# Excel styling + export
# =========================
def style_sheet_table(ws):
    header_fill = PatternFill("solid", fgColor=CABIFY["header"])
    header_font = Font(color=CABIFY["white"], bold=True)

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=ws.max_column):
        for cell in row:
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    ws.freeze_panes = "A2"

    for col_cells in ws.columns:
        col_letter = col_cells[0].column_letter
        max_len = 10
        for c in col_cells[:200]:
            v = "" if c.value is None else str(c.value)
            max_len = max(max_len, len(v))
        ws.column_dimensions[col_letter].width = min(max_len + 2, 45)

def build_excel(detalle: pd.DataFrame):
    wb = Workbook()
    wb.remove(wb.active)

    ws_l = wb.create_sheet("Listas")
    ws_l["A1"] = "Tipo_Incidencia"
    for i, v in enumerate(TIPO_OPTS, start=2):
        ws_l[f"A{i}"] = v
    ws_l["C1"] = "Clasificación Manual"
    for i, v in enumerate(CLASIF_OPTS, start=2):
        ws_l[f"C{i}"] = v
    ws_l.sheet_state = "hidden"

    ws = wb.create_sheet("Detalle")
    df_out = detalle.copy()
    df_out["Fecha"] = pd.to_datetime(df_out["Fecha"], errors="coerce")
    for r in dataframe_to_rows(df_out, index=False, header=True):
        ws.append(r)
    style_sheet_table(ws)

    # Fecha formato corto
    fecha_col = list(df_out.columns).index("Fecha") + 1
    for rr in range(2, ws.max_row + 1):
        ws.cell(rr, fecha_col).number_format = "DD/MM/YYYY"

    cols = list(df_out.columns)
    col_tipo = cols.index("Tipo_Incidencia") + 1
    col_clas = cols.index("Clasificación Manual") + 1

    dv_tipo = DataValidation(type="list", formula1="=Listas!$A$2:$A$4", allow_blank=False)
    dv_clas = DataValidation(type="list", formula1="=Listas!$C$2:$C$8", allow_blank=False)
    ws.add_data_validation(dv_tipo)
    ws.add_data_validation(dv_clas)

    dv_tipo.add(f"{ws.cell(2, col_tipo).coordinate}:{ws.cell(ws.max_row, col_tipo).coordinate}")
    dv_clas.add(f"{ws.cell(2, col_clas).coordinate}:{ws.cell(ws.max_row, col_clas).coordinate}")

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out


# =========================
# Sidebar
# =========================
with st.sidebar:
    st.header("Cargar archivos")
    f_asistencia = st.file_uploader("1) Asistencia (XLS/XLSX)", type=["xls", "xlsx"])
    f_inasist = st.file_uploader("2) Inasistencias (XLS/XLSX)", type=["xls", "xlsx"])
    f_planif = st.file_uploader("3) Planificación Turnos (CSV)", type=["csv"])
    f_codif = st.file_uploader("4) Codificación Turnos (CSV)", type=["csv"])

    st.divider()
    only_area = st.text_input("Filtrar Área (opcional)", value="AEROPUERTO")
    min_diff_h = st.number_input("Umbral diferencia horas (para incidencia)", value=0.5, step=0.25, min_value=0.0)

if not all([f_asistencia, f_inasist, f_planif, f_codif]):
    st.info("Sube los 4 archivos para comenzar.")
    st.stop()


# =========================
# Load
# =========================
try:
    df_asist, raw_asist = read_asistencia_b1b2(f_asistencia)
    df_inas, raw_inas = read_inasistencia_detect(f_inasist)
    df_plan = pd.read_csv(f_planif)
    df_cod = pd.read_csv(f_codif)
except Exception as e:
    st.error(str(e))
    st.stop()

df_plan.columns = [str(c).strip() for c in df_plan.columns]
df_cod.columns = [str(c).strip() for c in df_cod.columns]


# =========================
# DIAGNÓSTICO AUTOMÁTICO SI FALTA RUT EN ASISTENCIA
# =========================
c_rut_a = find_col(df_asist, ["RUT", "Rut", "R.U.T", "R.U.T."])
if c_rut_a is None:
    st.error("No encontré columna RUT en Asistencia. Abajo va el diagnóstico de lo que estoy leyendo.")
    with st.expander("Diagnóstico Asistencia (RAW filas 1-8 + columnas detectadas)", expanded=True):
        st.write("**RAW (sin header) - primeras 8 filas**")
        st.dataframe(raw_asist.head(8), use_container_width=True)

        st.write("**Columnas construidas (B1+B2, desde columna B)**")
        st.write(list(df_asist.columns))

        st.write("**Columnas normalizadas (para detectar caracteres raros)**")
        st.write([{c: _norm_colname(c)} for c in df_asist.columns[:80]])

        suspects = [c for c in df_asist.columns if "rut" in _norm_colname(c)]
        st.write("**Columnas sospechosas que contienen 'rut'**")
        st.write(suspects)

    st.stop()


# =========================
# Cargar columnas relevantes (Asistencia)
# =========================
c_area_a = find_col(df_asist, ["Área", "Area"])
c_fent = find_col(df_asist, ["Fecha Entrada", "FechaEntrada"])
c_hent = find_col(df_asist, ["Hora Entrada", "HoraEntrada"])
c_fsal = find_col(df_asist, ["Fecha Salida", "FechaSalida"])
c_hsal = find_col(df_asist, ["Hora Salida", "HoraSalida"])

c_rec_in = find_col(df_asist, ["Dentro del Recinto (Entrada)", "Dentro de Recinto(Entrada)"])
c_rec_out = find_col(df_asist, ["Dentro del Recinto (Salida)", "Dentro de Recinto(Salida)"])

c_nombre = find_col(df_asist, ["Nombre"])
c_pa = find_col(df_asist, ["Primer Apellido", "PrimerApellido"])
c_sa = find_col(df_asist, ["Segundo Apellido", "SegundoApellido"])
c_esp = find_col(df_asist, ["Especialidad"])
c_sup_a = find_col(df_asist, ["Supervisor"])
c_turno_txt = find_col(df_asist, ["Turno"])

if not c_fent:
    st.error("Asistencia: falta 'Fecha Entrada'. Revisa el diagnóstico (columnas detectadas).")
    st.stop()


# =========================
# Codificación (Sigla -> Horario)
# =========================
c_sigla = find_col(df_cod, ["Sigla"])
c_hor = find_col(df_cod, ["Horario"])
if not c_sigla or not c_hor:
    st.error("Codificación CSV debe tener columnas Sigla y Horario.")
    st.stop()

turno_to_hor = dict(
    zip(
        df_cod[c_sigla].astype(str).str.strip(),
        df_cod[c_hor].astype(str).str.strip(),
    )
)


# =========================
# Planificación (CSV ancho -> largo)
# =========================
col_name = find_col(df_plan, ["Nombre del Colaborador"])
col_rut = find_col(df_plan, ["RUT"])
col_area = find_col(df_plan, ["Área", "Area"])
col_sup = find_col(df_plan, ["Supervisor"])

for need, label in [(col_name, "Nombre del Colaborador"), (col_rut, "RUT"), (col_area, "Área"), (col_sup, "Supervisor")]:
    if not need:
        st.error(f"Planificación CSV: falta columna '{label}'.")
        st.stop()

df_plan["RUT_norm"] = df_plan[col_rut].apply(normalize_rut)
fixed = [col_name, col_rut, col_area, col_sup, "RUT_norm"]
date_cols = [c for c in df_plan.columns if c not in fixed]

df_pl_long = df_plan.melt(
    id_vars=fixed,
    value_vars=date_cols,
    var_name="Fecha",
    value_name="Turno_Cod",
)
df_pl_long["Fecha_dt"] = df_pl_long["Fecha"].apply(try_parse_date_any)
df_pl_long["Turno_Cod"] = df_pl_long["Turno_Cod"].astype(str).str.strip()
df_pl_long.loc[df_pl_long["Turno_Cod"].isin(["", "nan", "NaT", "None"]), "Turno_Cod"] = ""

if only_area and col_area in df_pl_long.columns:
    df_pl_long = df_pl_long[df_pl_long[col_area].astype(str).str.upper().str.contains(only_area.upper(), na=False)].copy()

min_dt = df_pl_long["Fecha_dt"].min()
max_dt = df_pl_long["Fecha_dt"].max()
if pd.isna(min_dt) or pd.isna(max_dt):
    st.error("Planificación: no pude interpretar fechas de encabezado.")
    st.stop()

c1, c2 = st.columns(2)
with c1:
    start_date = st.date_input("Desde", value=min_dt.date())
with c2:
    end_date = st.date_input("Hasta", value=max_dt.date())

if start_date > end_date:
    st.error("Rango inválido: Desde > Hasta")
    st.stop()

start_dt = pd.to_datetime(start_date)
end_dt = pd.to_datetime(end_date)
df_pl_long = df_pl_long[(df_pl_long["Fecha_dt"] >= start_dt) & (df_pl_long["Fecha_dt"] <= end_dt)].copy()

df_pl_long["Es_Libre"] = df_pl_long["Turno_Cod"].astype(str).str.upper().eq("L")
df_pl_act = df_pl_long[(df_pl_long["Turno_Cod"] != "") & (~df_pl_long["Es_Libre"])].copy()

df_pl_act["Horario_Plan"] = df_pl_act["Turno_Cod"].map(turno_to_hor).fillna("")
df_pl_act[["PlanStart_t", "PlanEnd_t"]] = df_pl_act["Horario_Plan"].apply(lambda x: pd.Series(parse_range_to_times(x)))

df_pl_act["Clave_RUT_Fecha"] = df_pl_act["RUT_norm"].astype(str) + "_" + df_pl_act["Fecha_dt"].dt.strftime("%Y-%m-%d").fillna("")
plan_lookup = df_pl_act.set_index("Clave_RUT_Fecha")[["Turno_Cod", "Horario_Plan", "PlanStart_t", "PlanEnd_t"]].to_dict(orient="index")
valid_ruts = set(df_pl_long["RUT_norm"].dropna().unique().tolist())


# =========================
# Filtrar Asistencia / Inasistencia por ruts válidos y fechas
# =========================
df_asist["RUT_norm"] = df_asist[c_rut_a].apply(normalize_rut)
df_asist = df_asist[df_asist["RUT_norm"].isin(valid_ruts)].copy()

if only_area and c_area_a:
    df_asist = df_asist[df_asist[c_area_a].astype(str).str.upper().str.contains(only_area.upper(), na=False)].copy()

df_asist["Fecha_base"] = df_asist[c_fent].apply(try_parse_date_any)
df_asist = df_asist[(df_asist["Fecha_base"] >= start_dt) & (df_asist["Fecha_base"] <= end_dt)].copy()
df_asist["Clave_RUT_Fecha"] = df_asist["RUT_norm"].astype(str) + "_" + df_asist["Fecha_base"].dt.strftime("%Y-%m-%d").fillna("")

c_rut_i = find_col(df_inas, ["RUT", "Rut", "R.U.T", "R.U.T."])
c_dia = find_col(df_inas, ["Día", "Dia"])
if c_rut_i is None or c_dia is None:
    st.error("Inasistencias: faltan columnas RUT y/o Día.")
    st.stop()

df_inas["RUT_norm"] = df_inas[c_rut_i].apply(normalize_rut)
df_inas = df_inas[df_inas["RUT_norm"].isin(valid_ruts)].copy()

df_inas["Fecha_base"] = df_inas[c_dia].apply(try_parse_date_any)
df_inas = df_inas[(df_inas["Fecha_base"] >= start_dt) & (df_inas["Fecha_base"] <= end_dt)].copy()
df_inas["Clave_RUT_Fecha"] = df_inas["RUT_norm"].astype(str) + "_" + df_inas["Fecha_base"].dt.strftime("%Y-%m-%d").fillna("")


# =========================
# Cálculos
# =========================
def planned_window_for(key: str):
    info = plan_lookup.get(key)
    if not info:
        return (None, None, "", "")
    return (info.get("PlanStart_t"), info.get("PlanEnd_t"), info.get("Turno_Cod", ""), info.get("Horario_Plan", ""))

def planned_hours_for(key: str):
    ps, pe, _, _ = planned_window_for(key)
    if ps is None or pe is None:
        return 0.0
    dt0 = datetime(2000, 1, 1, ps.hour, ps.minute, ps.second)
    dt1 = datetime(2000, 1, 1, pe.hour, pe.minute, pe.second)
    if dt1 < dt0:
        dt1 += timedelta(days=1)
    return round((dt1 - dt0).total_seconds() / 3600.0, 2)


# =========================
# Construcción Detalle (simple)
# =========================
mot_map = {"-": "Injustificada", "P": "Permiso", "L": "Licencia", "V": "Vacaciones", "C": "Compensado"}

as_rows = []
for _, r in df_asist.iterrows():
    key = r["Clave_RUT_Fecha"]
    ps, pe, _, plan_hor = planned_window_for(key)
    d0 = r["Fecha_base"].date() if not pd.isna(r["Fecha_base"]) else None

    plan_start = combine_date_time(d0, ps) if ps else None
    plan_end = combine_date_time(d0, pe) if pe else None
    plan_start, plan_end = ensure_overnight(plan_start, plan_end)

    d_in = try_parse_date_any(r.get(c_fent)).date() if c_fent else d0
    t_in = parse_time_only(r.get(c_hent)) if c_hent else None
    d_out = try_parse_date_any(r.get(c_fsal)).date() if c_fsal else d0
    t_out = parse_time_only(r.get(c_hsal)) if c_hsal else None

    real_start = combine_date_time(d_in, t_in) if (d_in and t_in) else None
    real_end = combine_date_time(d_out, t_out) if (d_out and t_out) else None
    real_start, real_end = ensure_overnight(real_start, real_end)

    horas_trab = round(hours_between(real_start, real_end), 2) if (real_start and real_end) else 0.0
    horas_plan = planned_hours_for(key)
    if horas_plan <= 0:
        continue

    diff = abs(horas_plan - horas_trab)
    if diff < float(min_diff_h):
        continue

    turno_marcado = "Sin Marca"
    if real_start and real_end:
        turno_marcado = f"{real_start.strftime('%H:%M')}-{real_end.strftime('%H:%M')}"

    rec_in = "Sin Marca" if not c_rec_in else (_clean_cell(r.get(c_rec_in)) or "Sin Marca")
    rec_out = "Sin Marca" if not c_rec_out else (_clean_cell(r.get(c_rec_out)) or "Sin Marca")

    as_rows.append({
        "Fecha": d0,
        "Nombre": "" if not c_nombre else _clean_cell(r.get(c_nombre)),
        "Primer Apellido": "" if not c_pa else _clean_cell(r.get(c_pa)),
        "Segundo Apellido": "" if not c_sa else _clean_cell(r.get(c_sa)),
        "RUT": _clean_cell(r.get(c_rut_a)),
        "Especialidad": "" if not c_esp else _clean_cell(r.get(c_esp)),
        "Turno": plan_hor if plan_hor else ("" if not c_turno_txt else _clean_cell(r.get(c_turno_txt))),
        "Supervisor": "" if not c_sup_a else _clean_cell(r.get(c_sup_a)),
        "Turno Marcado": turno_marcado,
        "Dentro del Recinto (Entrada)": rec_in,
        "Dentro del Recinto (Salida)": rec_out,
        "Detalle": f"HorasPlan={horas_plan} | HorasTrab={horas_trab} | Diff={round(diff,2)}",
        "Tipo_Incidencia": "Marcaje/Turno",
        "Clasificación Manual": "Seleccionar",
        "Minutos Retraso": 0,
        "Minutos Salida Anticipada": 0,
    })

in_rows = []
c_mot_i = find_col(df_inas, ["Motivo"])
c_nombre_i = find_col(df_inas, ["Nombre"])
c_pa_i = find_col(df_inas, ["Primer Apellido", "PrimerApellido"])
c_sa_i = find_col(df_inas, ["Segundo Apellido", "SegundoApellido"])
c_esp_i = find_col(df_inas, ["Especialidad"])
c_sup_i = find_col(df_inas, ["Supervisor"])
c_turno_i = find_col(df_inas, ["Turno"])

for _, r in df_inas.iterrows():
    key = r["Clave_RUT_Fecha"]
    d0 = r["Fecha_base"].date() if not pd.isna(r["Fecha_base"]) else None
    ps, pe, _, plan_hor = planned_window_for(key)

    mot = "" if not c_mot_i else _clean_cell(r.get(c_mot_i))
    clas = mot_map.get(mot.upper(), "Seleccionar")

    in_rows.append({
        "Fecha": d0,
        "Nombre": "" if not c_nombre_i else _clean_cell(r.get(c_nombre_i)),
        "Primer Apellido": "" if not c_pa_i else _clean_cell(r.get(c_pa_i)),
        "Segundo Apellido": "" if not c_sa_i else _clean_cell(r.get(c_sa_i)),
        "RUT": _clean_cell(r.get(c_rut_i)),
        "Especialidad": "" if not c_esp_i else _clean_cell(r.get(c_esp_i)),
        "Turno": plan_hor if plan_hor else ("" if not c_turno_i else _clean_cell(r.get(c_turno_i))),
        "Supervisor": "" if not c_sup_i else _clean_cell(r.get(c_sup_i)),
        "Turno Marcado": "Sin Marca",
        "Dentro del Recinto (Entrada)": "Sin Marca",
        "Dentro del Recinto (Salida)": "Sin Marca",
        "Detalle": f"Motivo={mot}",
        "Tipo_Incidencia": "Inasistencia",
        "Clasificación Manual": clas if clas else "Seleccionar",
        "Minutos Retraso": 0,
        "Minutos Salida Anticipada": 0,
    })

df_det = pd.concat([pd.DataFrame(as_rows), pd.DataFrame(in_rows)], ignore_index=True)
if df_det.empty:
    st.warning("No hay registros con los filtros actuales.")
    st.stop()

cols_order = [
    "Fecha", "Nombre", "Primer Apellido", "Segundo Apellido", "RUT",
    "Especialidad", "Turno", "Supervisor",
    "Turno Marcado", "Dentro del Recinto (Entrada)", "Dentro del Recinto (Salida)",
    "Detalle", "Tipo_Incidencia", "Clasificación Manual",
    "Minutos Retraso", "Minutos Salida Anticipada",
]
df_det = df_det[cols_order].sort_values(["Fecha", "RUT"], na_position="last").reset_index(drop=True)

st.subheader("Detalle para clasificar")
edited = st.data_editor(
    df_det,
    use_container_width=True,
    column_config={
        "Fecha": st.column_config.DateColumn(format="DD/MM/YYYY"),
        "Tipo_Incidencia": st.column_config.SelectboxColumn(options=TIPO_OPTS),
        "Clasificación Manual": st.column_config.SelectboxColumn(options=CLASIF_OPTS),
        "Minutos Retraso": st.column_config.NumberColumn(min_value=0, step=1),
        "Minutos Salida Anticipada": st.column_config.NumberColumn(min_value=0, step=1),
    },
)

st.subheader("Descarga")
excel_bytes = build_excel(edited)
st.download_button(
    "Descargar Excel (dropdowns)",
    data=excel_bytes,
    file_name="reporte_ausentismo_incidencias.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
