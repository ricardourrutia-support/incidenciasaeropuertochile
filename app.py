import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import io

# --- CONFIGURACIÓN Y ESTILOS ---
st.set_page_config(page_title="Gestión de Ausentismo", page_icon="✈️", layout="wide")

st.markdown("""
    <style>
    .block-container {padding-top: 1rem; padding-bottom: 5rem;}
    </style>
""", unsafe_allow_html=True)

# --- FUNCIONES DE SOPORTE ---

def cargar_datos(uploaded_file):
    """Carga archivos detectando CSV o Excel."""
    if uploaded_file is None: return None
    try:
        if uploaded_file.name.endswith('.csv'):
            try: return pd.read_csv(uploaded_file)
            except: return pd.read_csv(uploaded_file, sep=';')
        else: return pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"Error: {e}")
        return None

def parsear_turno(str_turno):
    """Convierte '20:00-07:00' en objetos time."""
    try:
        if pd.isna(str_turno) or '-' not in str(str_turno): return None, None
        parts = str(str_turno).split('-')
        return datetime.strptime(parts[0].strip(), "%H:%M").time(), datetime.strptime(parts[1].strip(), "%H:%M").time()
    except: return None, None

def calcular_minutos_exactos(row):
    """
    Calcula la diferencia exacta en minutos para pre-llenar la columna (9).
    Maneja cruce de medianoche.
    """
    try:
        turno = row.get('Turno')
        hora_real = row.get('Hora Entrada')
        fecha = row.get('Fecha Entrada')
        
        # Si no hay marcaje, retornamos 0 (o podría ser null)
        if pd.isna(hora_real) or str(hora_real).strip() == '-':
            return 0
        
        t_ini, t_fin = parsear_turno(turno)
        if t_ini is None: return 0

        # Parsear fechas y horas
        fecha_dt = pd.to_datetime(fecha, dayfirst=True).date()
        
        # Limpieza hora real
        h_str = str(hora_real).strip()
        try: h_real = datetime.strptime(h_str, "%H:%M:%S").time()
        except: h_real = datetime.strptime(h_str, "%H:%M").time()
        
        # Construcción Datetimes (Lógica Nocturna)
        dt_teorico = datetime.combine(fecha_dt, t_ini)
        
        # Caso especial: Si el turno es nocturno (ej 22:00) y la persona llega a las 00:15
        # La fecha del reporte de BUK suele ser la del inicio de turno. 
        # Si la hora real es mucho menor que la de inicio (ej 00 vs 22), asumimos día siguiente.
        if h_real < t_ini and (t_ini.hour - h_real.hour) > 12:
             dt_real = datetime.combine(fecha_dt + timedelta(days=1), h_real)
        else:
             dt_real = datetime.combine(fecha_dt, h_real)

        diff = (dt_real - dt_teorico).total_seconds() / 60
        
        # Retornamos solo si es retraso (positivo)
        if diff > 0:
            return int(diff)
        else:
            return 0 # Llegó temprano o a tiempo
    except:
        return 0

def formatear_detalle_entrada(row, tolerancia):
    """(4) Genera el texto legible."""
    hora_real = row.get('Hora Entrada')
    recinto = row.get('Dentro del Recinto (Entrada)', 'N/D')
    
    if pd.isna(hora_real) or str(hora_real).strip() == '-':
        return "Sin Marcaje [No registró entrada]"
    
    minutos_retraso = calcular_minutos_exactos(row)
    
    h_str = str(hora_real).strip()[:5] # Solo HH:MM
    
    if minutos_retraso > tolerancia:
        estado_texto = f"{minutos_retraso} min retraso"
    else:
        estado_texto = "OK"
        
    return f"Recinto: {recinto} | Hora: {h_str} [{estado_texto}]"

def formatear_detalle_salida(row):
    """(5) Genera el texto para salida."""
    hora_salida = row.get('Hora Salida')
    recinto = row.get('Dentro del Recinto (Salida)', 'N/D')
    if pd.isna(hora_salida) or str(hora_salida).strip() == '-':
        return f"Recinto: {recinto} | Hora: Sin Marca"
    return f"Recinto: {recinto} | Hora: {str(hora_salida)[:5]}"

def determinar_tipo_preliminar(row, tolerancia):
    """(6) Tipo Preliminar."""
    if pd.isna(row.get('Hora Entrada')) or str(row.get('Hora Entrada')).strip() == '-':
        return "Inasistencia"
    minutos = calcular_minutos_exactos(row)
    if minutos > tolerancia:
        return "Incidencia"
    return "OK"

# --- INTERFAZ PRINCIPAL ---

st.title("✈️ Plataforma de Gestión Operativa")
st.markdown("---")

with st.sidebar:
    st.header("1. Inputs")
    file_asist = st.file_uploader("Reporte Asistencia", type=["xls", "xlsx", "csv"])
    file_ina = st.file_uploader("Reporte Inasistencias", type=["xls", "xlsx", "csv"])
    
    st.divider()
    st.header("2. Reglas")
    tolerancia = st.number_input("Tolerancia (min)", value=15, min_value=0)

if file_asist and file_ina:
    # 1. Carga y ETL
    df_a = cargar_datos(file_asist)
    df_i = cargar_datos(file_ina)
    
    df_a.columns = df_a.columns.str.strip()
    df_i.columns = df_i.columns.str.strip()
    
    # Filtro Especialidad
    all_specs = sorted(list(set(df_a['Especialidad'].dropna()) | set(df_i['Especialidad'].dropna())))
    selected_specs = st.sidebar.multiselect("Filtrar Especialidad", all_specs, default=all_specs)
    
    df_a = df_a[df_a['Especialidad'].isin(selected_specs)].copy()
    df_i = df_i[df_i['Especialidad'].isin(selected_specs)].copy()

    # Preparar Datos
    df_a['Origen'] = 'Asistencia'
    df_i['Origen'] = 'Inasistencia'
    df_i['Fecha Entrada'] = df_i.get('Día', '')
    df_i['Hora Entrada'] = '-'
    df_i['Hora Salida'] = '-'
    
    cols_comunes = ['RUT', 'Nombre', 'Primer Apellido', 'Especialidad', 'Turno', 'Fecha Entrada', 'Hora Entrada', 'Hora Salida', 
                    'Dentro del Recinto (Entrada)', 'Dentro del Recinto (Salida)', 'Origen']
    
    for c in cols_comunes:
        if c not in df_a.columns: df_a[c] = None
        if c not in df_i.columns: df_i[c] = None
            
    df_master = pd.concat([df_a[cols_comunes], df_i[cols_comunes]], ignore_index=True)
    
    # --- CREACIÓN DE COLUMNAS (TUS 9 PUNTOS) ---
    
    df_master['Colaborador'] = df_master['Nombre'] + " " + df_master['Primer Apellido']
    df_master['Detalle Entrada'] = df_master.apply(lambda x: formatear_detalle_entrada(x, tolerancia), axis=1)
    df_master['Detalle Salida'] = df_master.apply(formatear_detalle_salida, axis=1)
    df_master['Tipo'] = df_master.apply(lambda x: determinar_tipo_preliminar(x, tolerancia), axis=1)
    
    # (9) Minutos Incidencia (Pre-cálculo)
    df_master['Minutos Incidencia'] = df_master.apply(calcular_minutos_exactos, axis=1)
    
    # Pre-clasificación
    def pre_clasificar(row):
        tipo = row['Tipo']
        detalle = row['Detalle Entrada']
        if tipo == 'Inasistencia': return 'Ausencia'
        if 'retraso' in detalle: return 'Retraso'
        return 'OK'

    df_master['Clasificación'] = df_master.apply(pre_clasificar, axis=1)
    df_master['Justificación'] = 'Injustificado' 

    # UI: Tabla
    columnas_finales = [
        'Colaborador', 'Fecha Entrada', 'Turno', 'Detalle Entrada', 'Detalle Salida', 
        'Tipo', 'Clasificación', 'Justificación', 'Minutos Incidencia'
    ]
    df_visual = df_master[columnas_finales].copy()

    st.subheader("📝 Detalle de Incidencias")
    st.caption("Edita la columna '(9) Minutos Incidencia' con el tiempo real a descontar.")

    edited_df = st.data_editor(
        df_visual,
        column_config={
            "Colaborador": st.column_config.TextColumn("1. Colaborador", disabled=True),
            "Fecha Entrada": st.column_config.TextColumn("2. Fecha", disabled=True),
            "Turno": st.column_config.TextColumn("3. Turno", disabled=True),
            "Detalle Entrada": st.column_config.TextColumn("4. Entrada", width="medium", disabled=True),
            "Detalle Salida": st.column_config.TextColumn("5. Salida", width="medium", disabled=True),
            "Tipo": st.column_config.TextColumn("6. Tipo", disabled=True),
            
            "Clasificación": st.column_config.SelectboxColumn(
                "7. Clasificación",
                options=["OK", "No Procede", "Ausencia", "Retraso", "Salida Anticipada", "Mixto"],
                required=True,
                width="small"
            ),
            "Justificación": st.column_config.SelectboxColumn(
                "8. Justificación",
                options=["Injustificado", "Justificado", "N.A."],
                required=True,
                width="small"
            ),
            
            # (9) COLUMNA NUEVA EDITABLE
            "Minutos Incidencia": st.column_config.NumberColumn(
                "9. Min. Real",
                help="Minutos de retraso/salida anticipada a descontar.",
                min_value=0,
                step=1,
                required=True
            )
        },
        hide_index=True,
        num_rows="fixed",
        use_container_width=True,
        height=600
    )

    # --- KPIs ---
    st.divider()
    st.subheader("📊 Reporte de KPIs")

    df_calc = edited_df[edited_df['Clasificación'] != 'No Procede'].copy()
    
    if len(df_calc) > 0:
        total_turnos = len(df_calc)
        
        # Total de minutos que el supervisor confirmó (solo de injustificados)
        minutos_descuento = df_calc[df_calc['Justificación'] == 'Injustificado']['Minutos Incidencia'].sum()
        
        # Problemas operativos (conteo)
        problemas = df_calc[
            (df_calc['Clasificación'].isin(['Ausencia', 'Retraso', 'Salida Anticipada', 'Mixto'])) & 
            (df_calc['Justificación'] == 'Injustificado')
        ]
        
        cumplimiento = ((total_turnos - len(problemas)) / total_turnos) * 100
        
        k1, k2, k3 = st.columns(3)
        k1.metric("Turnos Procesados", total_turnos)
        k2.metric("Minutos a Descontar", int(minutos_descuento), help="Suma de columna (9) donde Justificación es 'Injustificado'")
        k3.metric("Cumplimiento", f"{cumplimiento:.1f}%")
        
        # Exportar
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            edited_df.to_excel(writer, index=False, sheet_name='Detalle Validado')
            
        st.download_button("💾 Descargar Excel", buffer.getvalue(), "Reporte_Final.xlsx")

else:
    st.info("Carga los archivos en el menú lateral.")
