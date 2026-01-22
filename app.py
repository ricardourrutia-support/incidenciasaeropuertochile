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

# --- FUNCIONES DE CARGA INTELIGENTE ---

def detectar_encabezados_y_cargar(uploaded_file):
    """
    Busca automáticamente en qué fila están los títulos (RUT, Especialidad, etc.)
    y carga el archivo correctamente, sin importar si es CSV o Excel.
    """
    if uploaded_file is None: return None
    
    # 1. Determinar el motor de lectura según extensión
    es_csv = uploaded_file.name.endswith('.csv')
    
    try:
        # Leemos las primeras 10 filas sin encabezado para "escanear" el archivo
        if es_csv:
            try:
                df_preview = pd.read_csv(uploaded_file, header=None, nrows=10)
            except:
                uploaded_file.seek(0)
                df_preview = pd.read_csv(uploaded_file, header=None, nrows=10, sep=';')
        else:
            df_preview = pd.read_excel(uploaded_file, header=None, nrows=10)
        
        # 2. Buscar en qué fila están las palabras clave
        palabras_clave = ['RUT', 'ESPECIALIDAD', 'NOMBRE', 'TURNO']
        fila_header = 0 # Default
        encontrado = False
        
        for i, row in df_preview.iterrows():
            # Convertimos la fila a texto mayúscula para buscar
            fila_txt = " ".join(row.astype(str)).upper()
            
            # Si encontramos al menos 2 palabras clave, asumimos que esta es la cabecera
            coincidencias = sum(1 for p in palabras_clave if p in fila_txt)
            if coincidencias >= 2:
                fila_header = i
                encontrado = True
                break
        
        # 3. Recargar el archivo completo usando la fila detectada como header
        uploaded_file.seek(0) # Volver al inicio del archivo
        
        if es_csv:
            try:
                df_final = pd.read_csv(uploaded_file, header=fila_header)
            except:
                uploaded_file.seek(0)
                df_final = pd.read_csv(uploaded_file, header=fila_header, sep=';')
        else:
            df_final = pd.read_excel(uploaded_file, header=fila_header)
            
        return df_final, encontrado
        
    except Exception as e:
        st.error(f"Error procesando {uploaded_file.name}: {e}")
        return None, False

# --- FUNCIONES DE LÓGICA DE NEGOCIO ---

def parsear_turno(str_turno):
    """Convierte '20:00-07:00' en objetos time."""
    try:
        if pd.isna(str_turno) or '-' not in str(str_turno): return None, None
        parts = str(str_turno).split('-')
        return datetime.strptime(parts[0].strip(), "%H:%M").time(), datetime.strptime(parts[1].strip(), "%H:%M").time()
    except: return None, None

def calcular_minutos_exactos(row):
    """Calcula minutos de diferencia (retraso/salida anticipada)."""
    try:
        turno = row.get('Turno')
        hora_real = row.get('Hora Entrada')
        fecha = row.get('Fecha Entrada')
        
        if pd.isna(hora_real) or str(hora_real).strip() in ['-', '', 'nan']:
            return 0
        
        t_ini, t_fin = parsear_turno(turno)
        if t_ini is None: return 0

        # Parsear fecha
        fecha_dt = pd.to_datetime(fecha, dayfirst=True).date()
        
        # Parsear hora real
        h_str = str(hora_real).strip()
        try: h_real = datetime.strptime(h_str, "%H:%M:%S").time()
        except: 
            try: h_real = datetime.strptime(h_str, "%H:%M").time()
            except: return 0
        
        # Lógica Nocturna
        dt_teorico = datetime.combine(fecha_dt, t_ini)
        
        if h_real < t_ini and (t_ini.hour - h_real.hour) > 12:
             dt_real = datetime.combine(fecha_dt + timedelta(days=1), h_real)
        else:
             dt_real = datetime.combine(fecha_dt, h_real)

        diff = (dt_real - dt_teorico).total_seconds() / 60
        
        if diff > 0: return int(diff)
        else: return 0
    except:
        return 0

def formatear_detalle_entrada(row, tolerancia):
    """Genera texto legible entrada."""
    hora_real = row.get('Hora Entrada')
    recinto = row.get('Dentro del Recinto (Entrada)', 'N/D')
    
    if pd.isna(hora_real) or str(hora_real).strip() in ['-', '', 'nan']:
        return "Sin Marcaje [No registró entrada]"
    
    minutos = calcular_minutos_exactos(row)
    h_str = str(hora_real).strip()[:5]
    
    if minutos > tolerancia:
        estado = f"{minutos} min retraso"
    else:
        estado = "OK"
        
    return f"Recinto: {recinto} | Hora: {h_str} [{estado}]"

def formatear_detalle_salida(row):
    """Genera texto legible salida."""
    hora_salida = row.get('Hora Salida')
    recinto = row.get('Dentro del Recinto (Salida)', 'N/D')
    
    if pd.isna(hora_salida) or str(hora_salida).strip() in ['-', '', 'nan']:
        return f"Recinto: {recinto} | Hora: Sin Marca"
    return f"Recinto: {recinto} | Hora: {str(hora_salida)[:5]}"

def determinar_tipo_preliminar(row, tolerancia):
    """Tipo Preliminar."""
    if pd.isna(row.get('Hora Entrada')) or str(row.get('Hora Entrada')).strip() in ['-', '', 'nan']:
        return "Inasistencia"
    if calcular_minutos_exactos(row) > tolerancia:
        return "Incidencia"
    return "OK"

# --- INTERFAZ PRINCIPAL ---

st.title("✈️ Plataforma de Gestión Operativa")
st.markdown("---")

with st.sidebar:
    st.header("1. Inputs")
    st.info("El sistema detectará automáticamente dónde comienzan los datos.")
    
    file_asist = st.file_uploader("Reporte Asistencia", type=["xls", "xlsx", "csv"])
    file_ina = st.file_uploader("Reporte Inasistencias", type=["xls", "xlsx", "csv"])
    
    st.divider()
    st.header("2. Reglas")
    tolerancia = st.number_input("Tolerancia (min)", value=15, min_value=0)

if file_asist and file_ina:
    
    # 1. CARGA INTELIGENTE
    df_a, ok_a = detectar_encabezados_y_cargar(file_asist)
    df_i, ok_i = detectar_encabezados_y_cargar(file_ina)
    
    if ok_a and ok_i:
        # Normalización de columnas
        df_a.columns = df_a.columns.str.strip()
        df_i.columns = df_i.columns.str.strip()
        
        # Validar que exista la columna crítica 'Especialidad'
        if 'Especialidad' not in df_a.columns:
            st.error("Error: No se encontró la columna 'Especialidad' en el archivo de Asistencia. Verifica que el archivo contenga este encabezado.")
            st.write("Columnas detectadas:", df_a.columns.tolist())
        elif 'Especialidad' not in df_i.columns:
            st.error("Error: No se encontró la columna 'Especialidad' en el archivo de Inasistencias.")
        else:
            # 2. FILTRO ESPECIALIDAD
            all_specs = sorted(list(set(df_a['Especialidad'].dropna()) | set(df_i['Especialidad'].dropna())))
            selected_specs = st.sidebar.multiselect("Filtrar Especialidad", all_specs, default=all_specs)
            
            df_a = df_a[df_a['Especialidad'].isin(selected_specs)].copy()
            df_i = df_i[df_i['Especialidad'].isin(selected_specs)].copy()

            # 3. UNIFICACIÓN
            df_a['Origen'] = 'Asistencia'
            df_i['Origen'] = 'Inasistencia'
            
            # Mapeo BUK Inasistencia
            df_i['Fecha Entrada'] = df_i.get('Día', '') 
            df_i['Hora Entrada'] = '-'
            df_i['Hora Salida'] = '-'
            
            # Columnas maestras
            cols_comunes = ['RUT', 'Nombre', 'Primer Apellido', 'Especialidad', 'Turno', 
                            'Fecha Entrada', 'Hora Entrada', 'Hora Salida', 
                            'Dentro del Recinto (Entrada)', 'Dentro del Recinto (Salida)', 'Origen']
            
            for c in cols_comunes:
                if c not in df_a.columns: df_a[c] = None
                if c not in df_i.columns: df_i[c] = None
                    
            df_master = pd.concat([df_a[cols_comunes], df_i[cols_comunes]], ignore_index=True)

            # 4. CÁLCULOS (9 PUNTOS)
            df_master['Colaborador'] = df_master['Nombre'].astype(str) + " " + df_master['Primer Apellido'].astype(str)
            df_master['Minutos Incidencia'] = df_master.apply(calcular_minutos_exactos, axis=1)
            df_master['Detalle Entrada'] = df_master.apply(lambda x: formatear_detalle_entrada(x, tolerancia), axis=1)
            df_master['Detalle Salida'] = df_master.apply(formatear_detalle_salida, axis=1)
            df_master['Tipo'] = df_master.apply(lambda x: determinar_tipo_preliminar(x, tolerancia), axis=1)
            
            def pre_clasificar(row):
                tipo = row['Tipo']
                detalle = row['Detalle Entrada']
                if tipo == 'Inasistencia': return 'Ausencia'
                if 'retraso' in detalle: return 'Retraso'
                return 'OK'

            df_master['Clasificación'] = df_master.apply(pre_clasificar, axis=1)
            df_master['Justificación'] = 'Injustificado'

            # 5. VISUALIZACIÓN
            cols_finales = [
                'Colaborador', 'Fecha Entrada', 'Turno', 'Detalle Entrada', 'Detalle Salida',
                'Tipo', 'Clasificación', 'Justificación', 'Minutos Incidencia'
            ]
            
            df_visual = df_master[cols_finales].copy()
            
            st.subheader("📝 Detalle de Incidencias")
            edited_df = st.data_editor(
                df_visual,
                column_config={
                    "Colaborador": st.column_config.TextColumn("1. Colaborador", disabled=True),
                    "Fecha Entrada": st.column_config.TextColumn("2. Fecha", disabled=True),
                    "Turno": st.column_config.TextColumn("3. Turno", disabled=True),
                    "Detalle Entrada": st.column_config.TextColumn("4. Entrada", width="medium", disabled=True),
                    "Detalle Salida": st.column_config.TextColumn("5. Salida", width="medium", disabled=True),
                    "Tipo": st.column_config.TextColumn("6. Tipo", disabled=True),
                    "Clasificación": st.column_config.SelectboxColumn("7. Clasificación", options=["OK", "No Procede", "Ausencia", "Retraso", "Salida Anticipada", "Mixto"], required=True, width="small"),
                    "Justificación": st.column_config.SelectboxColumn("8. Justificación", options=["Injustificado", "Justificado", "N.A."], required=True, width="small"),
                    "Minutos Incidencia": st.column_config.NumberColumn("9. Min. Real", min_value=0, step=1, required=True)
                },
                hide_index=True, num_rows="fixed", use_container_width=True, height=600
            )

            # 6. KPI y EXPORTACIÓN
            st.divider()
            st.subheader("📊 Reporte de KPIs")
            df_calc = edited_df[edited_df['Clasificación'] != 'No Procede'].copy()
            
            if len(df_calc) > 0:
                total = len(df_calc)
                minutos_total = df_calc[df_calc['Justificación'] == 'Injustificado']['Minutos Incidencia'].sum()
                problemas = df_calc[(df_calc['Clasificación'] != 'OK') & (df_calc['Justificación'] == 'Injustificado')]
                cumplimiento = ((total - len(problemas)) / total) * 100
                
                k1, k2, k3 = st.columns(3)
                k1.metric("Turnos Procesados", total)
                k2.metric("Minutos Descuento", int(minutos_total))
                k3.metric("Cumplimiento", f"{cumplimiento:.1f}%")
                
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    edited_df.to_excel(writer, index=False, sheet_name='Detalle Validado')
                    df_calc[df_calc['Justificación'] == 'Injustificado'].to_excel(writer, index=False, sheet_name='Nómina')
                st.download_button("💾 Descargar Excel", buffer.getvalue(), "Reporte_Final.xlsx")
    else:
        st.error("No se pudieron detectar los encabezados. Asegúrate que los archivos contengan columnas como 'RUT', 'Especialidad' o 'Turno'.")

else:
    st.info("Carga los archivos en el menú lateral.")
