import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime, date, timedelta

from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation


APP_VERSION = "v2026-01-05_asistencia-header-entrada-salida"

st.set_page_config(page_title="Ausentismo e Incidencias Operativas", layout="wide")
st.title("Plataforma de Gestión de Ausentismo e Incidencias Operativas")
st.sidebar.success(f"APP RUNNING: {APP_VERSION}")


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
    s = s.replace("\u00A0", " ").replace("\u2007", " ").replace("\u202F", " ")
    s = s.strip().lower()
    s = (s.replace("á","a").replace("é","e").replace("í","i")
           .replace("ó","o").replace("ú","u").replace("ñ","n"))
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
    out, last = [], ""
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

def read_asistencia_grouped_2rows(file, sheet_name=0, drop_first_col_if_empty=True):
    """
    Asistencia con encabezado 2 filas y merges tipo:
      fila1: ... Entrada (merge sobre K-L), Salida (merge sobre M-N) ...
      fila2: ... Fecha/Hora bajo Entrada; Fecha/Hora bajo Salida ...

    Construye columnas:
      Entrada + Fecha -> Fecha Entrada
      Entrada + Hora  -> Hora Entrada
      Salida + Fecha  -> Fecha Salida
      Salida + Hora   -> Hora Salida
    """
    raw = _read_excel_raw_noheader(file, sheet_name=sheet_name)
    if len(raw) < 3:
        raise RuntimeError("Asistencia: el archivo tiene menos de 3 filas (necesito 2 de header + datos).")

    h1 = _ffill_row(raw.iloc[0].tolist())
    h2 = _ffill_row(raw.iloc[1].tolist())

    # A veces la columna A es basura vacía; si está vacía en header, la eliminamos.
    start_col = 0
    if drop_first_col_if_empty:
        if _clean_cell(h1[0]) == "" and _clean_cell(h2[0]) == "":
            start_col = 1

    h1 = h1[start_col:]
    h2 = h2[start_col:]

    cols = []
    for g, s in zip(h1, h2):
        g_clean = _clean_cell(g)
        s_clean = _clean_cell(s)

        g_norm = _norm_colname(g_clean)
        s_norm = _norm_colname(s_clean)

        # Mapeo especial Entrada/Salida + Fecha/Hora
        if g_norm == "entrada" and s_norm == "fecha":
            cols.append("Fecha Entrada")
        elif g_norm == "entrada" and s_norm == "hora":
            cols.append("Hora Entrada")
        elif g_norm == "salida" and s_norm == "fecha":
            cols.append("Fecha Salida")
        elif g_norm == "salida" and s_norm == "hora":
            cols.append("Hora Salida")
        else:
            if g_clean and s_clean and _norm_colname(g_clean) != _norm_colname(s_clean):
                cols.append(f"{g_clean} {s_clean}".strip())
            elif s_clean:
                cols.append(s_clean)
            else:
                cols.append(g_clean)

    cols = [c if c else f"COL_{i+1}" for i, c in enumerate(cols)]
    cols = [_clean_cell(c) for c in cols]

    df = raw.iloc[2:, start_col:].copy()
    df.columns = cols
    df = df.dropna(how="all")
    df.columns = [_clean_cell(c) for c in df.columns]
    return df, raw

def read_inasistencia_detect(file, sheet_name=0, max_scan_rows=120):
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
# Excel export (dropdowns)
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
    if "Fecha" in df_out.columns:
        df_out["Fecha"] = pd.to_datetime(df_out["Fecha"], errors="coerce")

    for r in dataframe_to_rows(df_out, index=False, header=True):
        ws.append(r)
    style_sheet_table(ws)

    if "Fecha" in df_out.columns:
        fecha_col = list(df_out.columns).index("Fecha") + 1
        for rr in range(2, ws.max_row + 1):
            ws.cell(rr, fecha_col).number_format = "DD/MM/YYYY"

    cols = list(df_out.columns)
    if "Tipo_Incidencia" in cols and "Clasificación Manual" in cols:
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
    df_asist, raw_asist = read_asistencia_grouped_2rows(f_asistencia)
    df_inas, raw_inas = read_inasistencia_detect(f_inasist)
    df_plan = pd.read_csv(f_planif)
    df_cod = pd.read_csv(f_codif)
except Exception as e:
    st.error(str(e))
    st.stop()

df_plan.columns = [str(c).strip() for c in df_plan.columns]
df_cod.columns = [str(c).strip() for c in df_cod.columns]


# =========================
# DIAGNÓSTICO (VISIBLE SIEMPRE)
# =========================
with st.expander("Diagnóstico de lectura (Asistencia / Inasistencias)", expanded=True):
    st.write("**Asistencia: columnas detectadas**")
    st.write(list(df_asist.columns))
    st.write("**Asistencia: RAW primeras 6 filas (sin header)**")
    st.dataframe(raw_asist.head(6), use_container_width=True)
    st.write("**Asistencia: primeras 5 filas ya interpretadas**")
    st.dataframe(df_asist.head(5), use_container_width=True)

    st.write("---")
    st.write("**Inasistencias: columnas detectadas**")
    st.write(list(df_inas.columns))
    st.write("**Inasistencias: RAW primeras 6 filas (sin header)**")
    st.dataframe(raw_inas.head(6), use_container_width=True)
    st.write("**Inasistencias: primeras 5 filas ya interpretadas**")
    st.dataframe(df_inas.head(5), use_container_width=True)


# =========================
# Validaciones mínimas de Asistencia
# =========================
c_rut_a = find_col(df_asist, ["RUT", "Rut", "R.U.T", "R.U.T."])
if c_rut_a is None:
    suspects = [c for c in df_asist.columns if "rut" in _norm_colname(c)]
    st.error(f"No encontré columna RUT en Asistencia. Columnas sospechosas: {suspects}")
    st.stop()

c_fent = find_col(df_asist, ["Fecha Entrada"])
c_hent = find_col(df_asist, ["Hora Entrada"])
c_fsal = find_col(df_asist, ["Fecha Salida"])
c_hsal = find_col(df_asist, ["Hora Salida"])

if c_fent is None:
    st.error("Asistencia: falta 'Fecha Entrada' (debería construirse desde Entrada+Fecha). Revisa el diagnóstico arriba.")
    st.stop()


# =========================
# (Resto del flujo)
# =========================
# Por ahora dejamos una salida mínima para confirmar que el header ya está bien.
st.success("✅ Asistencia leída correctamente: encontré RUT y Fecha Entrada / Hora Entrada / Fecha Salida / Hora Salida.")

st.write("Columnas clave detectadas:")
st.write({
    "RUT": c_rut_a,
    "Fecha Entrada": c_fent,
    "Hora Entrada": c_hent,
    "Fecha Salida": c_fsal,
    "Hora Salida": c_hsal,
})

st.subheader("Descarga (prueba)")
# Export simple para que pruebes que abre y que los dropdown funcionan (puedes conectar el resto después).
demo = pd.DataFrame({
    "Fecha": pd.to_datetime(df_asist[c_fent], errors="coerce", dayfirst=True),
    "RUT": df_asist[c_rut_a].astype(str),
    "Tipo_Incidencia": "Marcaje/Turno",
    "Clasificación Manual": "Seleccionar",
})
excel_bytes = build_excel(demo)
st.download_button(
    "Descargar Excel (prueba dropdowns)",
    data=excel_bytes,
    file_name="diagnostico_asistencia.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
