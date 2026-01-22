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

# --- FUNCIONES DE NORMALIZACIÓN (NUEVO MOTOR) ---

def normalizar_nombres_cols(df):
    """
    Renombra las columnas del DataFrame a un estándar común,
    buscando coincidencias flexibles (ej. 'Fecha' -> 'Fecha Entrada').
    """
    if df is None or df.empty:
        return df

    # Mapa de búsqueda: {Nombre_Estandar: [Lista de Posibles Nombres en BUK]}
    mapa = {
        'Fecha Entrada': ['fecha entrada', 'fecha', 'dia', 'día', 'date'],
        'Hora Entrada': ['hora entrada', 'entrada', 'hora inicio', 'inicio', 'marcado entrada'],
        'Hora Salida': ['hora salida', 'salida', 'hora fin', 'fin', 'marcado salida'],
        'Turno': ['turno', 'jornada', 'horario'],
        'Especialidad': ['especialidad', 'cargo', 'puesto', 'job'],
        'Nombre': ['nombre', 'nombres', 'name'],
        'Primer Apellido': ['primer apellido', 'apellido paterno', 'apellido 1', 'last name'],
        'RUT': ['rut', 'dni', 'identificación'],
        'Dentro del Recinto (Entrada)': ['dentro del recinto (entrada)', 'recinto entrada', 'zona entrada'],
        'Dentro del Recinto (Salida)': ['dentro del recinto (salida)', 'recinto salida', 'zona salida']
    }

    # Crear diccionario de renombrado
    nuevo_cols = {}
    cols_existentes = {c.lower().strip(): c for c in df.columns} # Mapa minuscula -> nombre real

    for estandar, posibles in mapa.items():
        for candidato in posibles:
            candidato_clean = candidato.lower().strip()
            # Buscar coincidencia exacta o parcial
            matches = [real for k, real in cols_existentes.items() if candidato_clean == k]
            
            if matches:
                nuevo_cols[matches[0]] = estandar
                break # Encontramos una, pasamos a la siguiente columna estándar

    # Aplicar renombrado
    df_renombrado = df.rename(columns=nuevo_cols)
    return df_renombrado

def detectar_encabezados_y_cargar(uploaded_file):
    """Carga inteligente detectando la fila de cabecera."""
    if uploaded_file is None: return None
    
    es_csv = uploaded_file.name.endswith('.csv')
    
    try:
        # Pre-lectura
        if es_csv:
            try: df_preview = pd.read_csv(uploaded_file, header=None, nrows=10)
            except: 
                uploaded_file.seek(0)
                df_preview = pd.read_csv(uploaded_file, header=None, nrows=10, sep=';')
        else:
            df_preview = pd.read_excel(uploaded_file, header=None, nrows=10)
        
        # Buscar fila con palabras clave
        palabras_clave = ['RUT', 'TURNO', 'NOMBRE'] # Palabras muy comunes
        fila_header = 0
        encontrado = False
        
        for i, row in df_preview.iterrows():
            fila_txt = " ".join(row.astype(str)).upper()
            if sum(1 for p in palabras_clave if p in fila_txt) >= 2:
                fila_header = i
                encontrado = True
                break
        
        # Carga real
        uploaded_file.seek(0)
        if es_csv:
            try: df_final = pd.read_csv(uploaded_file, header=fila_header)
            except: 
                uploaded_file.seek(0)
                df_final = pd.read_csv(uploaded_file, header=fila_header, sep=';')
        else:
            df_final = pd.read_excel(uploaded_file, header=fila_header)
            
        return df_final, encontrado
        
    except Exception as e:
        st.error(f"Error procesando archivo: {e}")
        return None, False

# --- FUNCIONES LÓGICA NEGOCIO ---

def parsear_turno(str_turno):
    try:
        if pd.isna(str_turno) or '-' not in str(str_turno): return None, None
        parts = str(str_turno).split('-')
        return datetime.strptime(parts[0].strip(), "%H:%M").time(), datetime.strptime(parts[1].strip(), "%H:%M").time()
    except: return None, None

def calcular_minutos_exactos(row):
    try:
        turno = row.get('Turno')
        hora_real = row.get('Hora Entrada')
        fecha = row.get('Fecha Entrada')
        
        if pd.isna(hora_real) or str(hora_real).strip() in ['-', '', 'nan', 'NaT']: return 0
        
        t_ini, t_fin = parsear_turno(turno)
        if t_ini is None: return 0

        # Fecha
        if isinstance(fecha, str):
            fecha_dt = pd.to_datetime(fecha, dayfirst=True).date()
        else:
            fecha_dt = fecha.date() if hasattr(fecha, 'date') else fecha

        # Hora
        h_str = str(hora_real).strip()
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

def formatear_detalle_entrada(row, tolerancia):
    hora_real = row.get('Hora Entrada')
    recinto = row.get('Dentro del Recinto (Entrada)', 'N/D')
    
    if pd.isna(hora_real) or str(hora_real).strip() in ['-', '', 'nan']:
        return "Sin Marcaje [No registró entrada]"
    
    minutos = calcular_minutos_exactos(row)
    h_str = str(hora_real).strip()[:5]
    
    estado = f"{minutos} min retraso" if minutos > tolerancia else "OK"
    return f"Recinto: {recinto} | Hora: {h_str} [{estado}]"

def formatear_detalle_salida(row):
    hora_salida = row.get('Hora Salida')
    recinto = row.get('Dentro del Recinto (Salida)', 'N/D')
    if pd.isna(hora_salida) or str(hora_salida).strip() in ['-', '', 'nan']:
        return f"Recinto: {recinto} | Hora: Sin Marca"
    return f"Recinto: {recinto} | Hora: {str(hora_salida)[:5]}"

def determinar_tipo_preliminar(row, tolerancia):
    if pd.isna(row.get('Hora Entrada')) or str(row.get('Hora Entrada')).strip() in ['-', '', 'nan']:
        return "Inasistencia"
    if calcular_minutos_exactos(row) > tolerancia:
        return "Incidencia"
    return "OK"

# --- MAIN ---

st.title("✈️ Plataforma de Gestión Operativa")
st.markdown("---")

with st.sidebar:
    st.header("1. Inputs")
    file_asist = st.file_uploader("Reporte Asistencia", type=["xls", "xlsx", "csv"])
    file_ina = st.file_uploader("Reporte Inasistencias", type=["xls", "xlsx", "csv"])
    st.divider()
    tolerancia = st.number_input("Tolerancia (min)", value=15, min_value=0)

if file_asist and file_ina:
    # 1. Carga
    df_a, ok_a = detectar_encabezados_y_cargar(file_asist)
    df_i, ok_i = detectar_encabezados_y_cargar(file_ina)
    
    if ok_a and ok_i:
        # 2. NORMALIZACIÓN DE COLUMNAS (Aquí ocurre la magia)
        df_a = normalizar_nombres_cols(df_a)
        df_i = normalizar_nombres_cols(df_i)
        
        # Validar columnas críticas post-normalización
        req_cols = ['Especialidad', 'Turno'] # Mínimo necesario
        missing_a = [c for c in req_cols if c not in df_a.columns]
        
        if missing_a:
            st.error(f"Faltan columnas clave en Asistencia: {missing_a}. Columnas encontradas: {df_a.columns.tolist()}")
        else:
            # 3. Filtros
            all_specs = sorted(list(set(df_a['Especialidad'].dropna()) | set(df_i['Especialidad'].dropna())))
            selected_specs = st.sidebar.multiselect("Filtrar Especialidad", all_specs, default=all_specs)
            
            df_a = df_a[df_a['Especialidad'].isin(selected_specs)].copy()
            df_i = df_i[df_i['Especialidad'].isin(selected_specs)].copy()

            # 4. Unificación
            df_a['Origen'] = 'Asistencia'
            df_i['Origen'] = 'Inasistencia'
            
            # Si inasistencia no tiene Fecha Entrada pero tiene Día (por si falló normalización)
            if 'Fecha Entrada' not in df_i.columns and 'Día' in df_i.columns:
                 df_i['Fecha Entrada'] = df_i['Día']
            
            # Asegurar columnas vacías
            cols_std = ['RUT', 'Nombre', 'Primer Apellido', 'Especialidad', 'Turno', 
                        'Fecha Entrada', 'Hora Entrada', 'Hora Salida', 
                        'Dentro del Recinto (Entrada)', 'Dentro del Recinto (Salida)', 'Origen']
            
            for c in cols_std:
                if c not in df_a.columns: df_a[c] = None
                if c not in df_i.columns: df_i[c] = None
            
            df_master = pd.concat([df_a[cols_std], df_i[cols_std]], ignore_index=True)

            # 5. Cálculos
            df_master['Colaborador'] = df_master['Nombre'].astype(str) + " " + df_master['Primer Apellido'].astype(str)
            df_master['Minutos Incidencia'] = df_master.apply(calcular_minutos_exactos, axis=1)
            df_master['Detalle Entrada'] = df_master.apply(lambda x: formatear_detalle_entrada(x, tolerancia), axis=1)
            df_master['Detalle Salida'] = df_master.apply(formatear_detalle_salida, axis=1)
            df_master['Tipo'] = df_master.apply(lambda x: determinar_tipo_preliminar(x, tolerancia), axis=1)
            
            def pre_clasificar(row):
                if row['Tipo'] == 'Inasistencia': return 'Ausencia'
                if 'retraso' in row['Detalle Entrada']: return 'Retraso'
                return 'OK'

            df_master['Clasificación'] = df_master.apply(pre_clasificar, axis=1)
            df_master['Justificación'] = 'Injustificado'

            # 6. Visualización
            cols_ui = ['Colaborador', 'Fecha Entrada', 'Turno', 'Detalle Entrada', 'Detalle Salida',
                       'Tipo', 'Clasificación', 'Justificación', 'Minutos Incidencia']
            
            st.subheader("📝 Detalle")
            edited_df = st.data_editor(
                df_master[cols_ui],
                column_config={
                    "Colaborador": st.column_config.TextColumn(disabled=True),
                    "Fecha Entrada": st.column_config.TextColumn(disabled=True),
                    "Turno": st.column_config.TextColumn(disabled=True),
                    "Detalle Entrada": st.column_config.TextColumn(disabled=True, width="medium"),
                    "Detalle Salida": st.column_config.TextColumn(disabled=True, width="medium"),
                    "Tipo": st.column_config.TextColumn(disabled=True),
                    "Clasificación": st.column_config.SelectboxColumn(options=["OK", "No Procede", "Ausencia", "Retraso", "Salida Anticipada", "Mixto"], required=True),
                    "Justificación": st.column_config.SelectboxColumn(options=["Injustificado", "Justificado", "N.A."], required=True),
                    "Minutos Incidencia": st.column_config.NumberColumn(min_value=0, step=1, required=True)
                },
                hide_index=True, use_container_width=True, height=600
            )

            # 7. KPI y Export
            st.divider()
            st.subheader("📊 KPIs")
            df_kpi = edited_df[edited_df['Clasificación'] != 'No Procede']
            
            if not df_kpi.empty:
                total = len(df_kpi)
                mins = df_kpi[df_kpi['Justificación'] == 'Injustificado']['Minutos Incidencia'].sum()
                probs = len(df_kpi[(df_kpi['Clasificación'] != 'OK') & (df_kpi['Justificación'] == 'Injustificado')])
                cumplimiento = ((total - probs) / total) * 100
                
                c1, c2, c3 = st.columns(3)
                c1.metric("Turnos", total)
                c2.metric("Minutos Descuento", int(mins))
                c3.metric("Cumplimiento", f"{cumplimiento:.1f}%")
                
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    edited_df.to_excel(writer, index=False, sheet_name='Detalle')
                    df_kpi[df_kpi['Justificación'] == 'Injustificado'].to_excel(writer, index=False, sheet_name='Nomina')
                st.download_button("💾 Excel", buffer.getvalue(), "Reporte.xlsx")
    else:
        st.error("No se detectaron encabezados válidos.")
else:
    st.info("Sube los archivos.")
