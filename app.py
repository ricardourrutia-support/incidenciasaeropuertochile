import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import io

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="Gestión de Ausentismo", page_icon="✈️", layout="wide")

st.markdown("""
    <style>
    .block-container {padding-top: 1rem; padding-bottom: 5rem;}
    .stAlert {padding: 0.5rem;}
    </style>
""", unsafe_allow_html=True)

# --- FUNCIONES DE CARGA ROBUSTA (SOLUCIÓN DEFINITIVA) ---

def reparar_encabezados_buk(df):
    """
    Detecta y repara el problema específico de las celdas combinadas de BUK.
    Convierte: 'Entrada' -> 'Fecha Entrada', 'Unnamed: 11' -> 'Hora Entrada'.
    """
    columnas_nuevas = {}
    cols = df.columns.tolist()
    
    # 1. Mapeo Posicional (Lo más seguro para este reporte)
    # Buscamos patrones conocidos en los nombres que arroja pandas al leer row 0
    for i, col in enumerate(cols):
        c_clean = str(col).strip().lower()
        
        # Caso Columna 'Entrada' (que en realidad es la Fecha)
        if c_clean == 'entrada':
            columnas_nuevas[col] = 'Fecha Entrada'
            # La columna siguiente suele ser la Hora (que viene como Unnamed o NaN)
            if i + 1 < len(cols):
                columnas_nuevas[cols[i+1]] = 'Hora Entrada'
        
        # Caso Columna 'Salida'
        elif c_clean == 'salida':
            columnas_nuevas[col] = 'Fecha Salida'
            if i + 1 < len(cols):
                columnas_nuevas[cols[i+1]] = 'Hora Salida'
                
        # Caso RUT Empleador vs RUT Colaborador
        elif 'rut' in c_clean and 'empleador' not in c_clean:
            columnas_nuevas[col] = 'RUT'
            
        elif 'dentro de recinto' in c_clean and 'entrada' in c_clean:
            columnas_nuevas[col] = 'Dentro del Recinto (Entrada)'
            
        elif 'dentro de recinto' in c_clean and 'salida' in c_clean:
            columnas_nuevas[col] = 'Dentro del Recinto (Salida)'

    # Aplicar renombrado detectado
    if columnas_nuevas:
        df = df.rename(columns=columnas_nuevas)
    
    return df

def cargar_datos_inteligente(uploaded_file):
    if uploaded_file is None: return None, False
    
    try:
        # 1. Determinar formato y leer primeras filas para inspección
        es_csv = uploaded_file.name.endswith('.csv')
        
        if es_csv:
            try: df_raw = pd.read_csv(uploaded_file, header=None, nrows=15)
            except: 
                uploaded_file.seek(0)
                df_raw = pd.read_csv(uploaded_file, header=None, nrows=15, sep=';')
        else:
            df_raw = pd.read_excel(uploaded_file, header=None, nrows=15)

        # 2. Buscar la fila REAL de encabezados
        # Buscamos donde aparezcan "RUT" y "TURNO" en la misma fila
        fila_header = -1
        for i, row in df_raw.iterrows():
            fila_txt = " ".join(row.astype(str)).upper()
            if "RUT" in fila_txt and "TURNO" in fila_txt:
                fila_header = i
                break
        
        if fila_header == -1:
            st.error(f"No se detectaron encabezados válidos en {uploaded_file.name}. Buscando 'RUT' y 'TURNO'.")
            return None, False

        # 3. Cargar archivo desde esa fila
        uploaded_file.seek(0)
        if es_csv:
            try: df = pd.read_csv(uploaded_file, header=fila_header)
            except: 
                uploaded_file.seek(0)
                df = pd.read_csv(uploaded_file, header=fila_header, sep=';')
        else:
            df = pd.read_excel(uploaded_file, header=fila_header)

        # 4. REPARACIÓN CRÍTICA BUK
        # Si vemos columnas llamadas "Entrada" y "Unnamed", aplicamos el parche
        df = reparar_encabezados_buk(df)
        
        # 5. Normalización adicional (por seguridad)
        mapa_extra = {
            'Especialidad': ['cargo', 'puesto'],
            'Nombre': ['nombres'],
            'Turno': ['jornada', 'horario']
        }
        cols_map = {}
        for real_col in df.columns:
            for estandar, variantes in mapa_extra.items():
                if real_col.lower().strip() in variantes:
                    cols_map[real_col] = estandar
        df = df.rename(columns=cols_map)
        
        return df, True

    except Exception as e:
        st.error(f"Error procesando {uploaded_file.name}: {e}")
        return None, False

# --- FUNCIONES DE CÁLCULO ---

def parsear_turno(str_turno):
    try:
        if pd.isna(str_turno) or '-' not in str(str_turno): return None, None
        parts = str(str_turno).split('-')
        return datetime.strptime(parts[0].strip(), "%H:%M").time(), datetime.strptime(parts[1].strip(), "%H:%M").time()
    except: return None, None

def calcular_minutos_exactos(row):
    try:
        # Extraer datos usando nombres normalizados
        fecha = row.get('Fecha Entrada')
        hora = row.get('Hora Entrada')
        turno = row.get('Turno')

        if pd.isna(hora) or str(hora).strip() in ['-', '', 'nan', 'NaT']: return 0
        
        t_ini, t_fin = parsear_turno(turno)
        if t_ini is None: return 0
        
        # Parsear Fecha
        if isinstance(fecha, str):
            fecha_dt = pd.to_datetime(fecha, dayfirst=True).date()
        else:
            fecha_dt = fecha.date() if hasattr(fecha, 'date') else fecha

        # Parsear Hora
        h_str = str(hora).strip()
        try: h_real = datetime.strptime(h_str, "%H:%M:%S").time()
        except: 
            try: h_real = datetime.strptime(h_str, "%H:%M").time()
            except: return 0
            
        dt_teorico = datetime.combine(fecha_dt, t_ini)
        
        # Lógica Nocturna
        if h_real < t_ini and (t_ini.hour - h_real.hour) > 12:
             dt_real = datetime.combine(fecha_dt + timedelta(days=1), h_real)
        else:
             dt_real = datetime.combine(fecha_dt, h_real)

        diff = (dt_real - dt_teorico).total_seconds() / 60
        return int(diff) if diff > 0 else 0
    except:
        return 0

def formatear_detalle(row, tipo='entrada', tolerancia=0):
    col_fecha = 'Fecha Entrada' if tipo == 'entrada' else 'Fecha Salida'
    col_hora = 'Hora Entrada' if tipo == 'entrada' else 'Hora Salida'
    col_recinto = 'Dentro del Recinto (Entrada)' if tipo == 'entrada' else 'Dentro del Recinto (Salida)'
    
    hora = row.get(col_hora)
    recinto = row.get(col_recinto, 'N/D')
    
    if pd.isna(hora) or str(hora).strip() in ['-', '', 'nan']:
        return "Sin Marcaje"
    
    h_str = str(hora).strip()[:5]
    
    if tipo == 'entrada':
        mins = calcular_minutos_exactos(row)
        estado = f"{mins} min retraso" if mins > tolerancia else "OK"
        return f"Recinto: {recinto} | Hora: {h_str} [{estado}]"
    else:
        return f"Recinto: {recinto} | Hora: {h_str}"

# --- INTERFAZ ---

st.title("✈️ Plataforma de Gestión Operativa (BUK Fixed)")

with st.sidebar:
    st.header("1. Carga de Archivos")
    f_asist = st.file_uploader("Asistencia", type=["xls", "xlsx", "csv"])
    f_ina = st.file_uploader("Inasistencias", type=["xls", "xlsx", "csv"])
    st.divider()
    tolerancia = st.number_input("Tolerancia (min)", 15)

if f_asist and f_ina:
    df_a, ok_a = cargar_datos_inteligente(f_asist)
    df_i, ok_i = cargar_datos_inteligente(f_ina)
    
    if ok_a and ok_i:
        # Filtros
        specs = sorted(list(set(df_a['Especialidad'].dropna()) | set(df_i['Especialidad'].dropna())))
        sel_specs = st.sidebar.multiselect("Especialidad", specs, default=specs)
        
        df_a = df_a[df_a['Especialidad'].isin(sel_specs)].copy()
        df_i = df_i[df_i['Especialidad'].isin(sel_specs)].copy()
        
        # Preparar Inasistencias
        df_i['Origen'] = 'Inasistencia'
        if 'Día' in df_i.columns: df_i['Fecha Entrada'] = df_i['Día']
        
        # Unificar
        cols = ['RUT', 'Nombre', 'Primer Apellido', 'Especialidad', 'Turno', 
                'Fecha Entrada', 'Hora Entrada', 'Fecha Salida', 'Hora Salida',
                'Dentro del Recinto (Entrada)', 'Dentro del Recinto (Salida)']
        
        # Asegurar columnas
        for c in cols:
            if c not in df_a.columns: df_a[c] = None
            if c not in df_i.columns: df_i[c] = None
            
        df_master = pd.concat([df_a, df_i], ignore_index=True)
        
        # Cálculos Finales
        df_master['Colaborador'] = df_master['Nombre'].astype(str) + " " + df_master['Primer Apellido'].astype(str)
        df_master['Minutos Incidencia'] = df_master.apply(calcular_minutos_exactos, axis=1)
        df_master['Detalle Entrada'] = df_master.apply(lambda x: formatear_detalle(x, 'entrada', tolerancia), axis=1)
        df_master['Detalle Salida'] = df_master.apply(lambda x: formatear_detalle(x, 'salida'), axis=1)
        
        # Clasificación
        def clasificar(row):
            if pd.isna(row.get('Hora Entrada')) and row.get('Origen') == 'Inasistencia': return 'Ausencia'
            if pd.isna(row.get('Hora Entrada')): return 'Ausencia' # Caso asistencia vacía
            if row['Minutos Incidencia'] > tolerancia: return 'Retraso'
            return 'OK'
            
        df_master['Clasificación'] = df_master.apply(clasificar, axis=1)
        df_master['Justificación'] = 'Injustificado'
        
        # Tabla
        cols_view = ['Colaborador', 'Fecha Entrada', 'Turno', 'Detalle Entrada', 'Detalle Salida', 
                     'Clasificación', 'Justificación', 'Minutos Incidencia']
        
        st.subheader("📝 Validación")
        edited = st.data_editor(
            df_master[cols_view],
            column_config={
                "Colaborador": st.column_config.TextColumn(disabled=True),
                "Fecha Entrada": st.column_config.TextColumn(disabled=True),
                "Detalle Entrada": st.column_config.TextColumn(width="medium", disabled=True),
                "Clasificación": st.column_config.SelectboxColumn(options=["OK", "Ausencia", "Retraso", "No Procede"], required=True),
                "Justificación": st.column_config.SelectboxColumn(options=["Injustificado", "Justificado"], required=True),
                "Minutos Incidencia": st.column_config.NumberColumn(min_value=0, step=1)
            },
            use_container_width=True, height=600, hide_index=True
        )
        
        # Exportar
        st.divider()
        if st.button("Generar Reporte Excel"):
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                edited.to_excel(writer, index=False)
            st.download_button("📥 Descargar", buffer.getvalue(), "Reporte_Final.xlsx")

    else:
        st.error("Error al leer archivos. Verifica el formato.")
else:
    st.info("Sube los archivos en el menú lateral.")
