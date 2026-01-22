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

def cargar_datos(uploaded_file, saltar_filas=0):
    """
    Carga archivos detectando CSV o Excel.
    Recibe 'saltar_filas' para manejar encabezados complejos (ej. Asistencia empieza en fila 3).
    """
    if uploaded_file is None: return None
    try:
        if uploaded_file.name.endswith('.csv'):
            try: 
                return pd.read_csv(uploaded_file, skiprows=saltar_filas)
            except: 
                return pd.read_csv(uploaded_file, sep=';', skiprows=saltar_filas)
        else:
            # Para Excel
            return pd.read_excel(uploaded_file, skiprows=saltar_filas)
    except Exception as e:
        st.error(f"Error cargando {uploaded_file.name}: {e}")
        return None

def parsear_turno(str_turno):
    """Convierte '20:00-07:00' en objetos time."""
    try:
        if pd.isna(str_turno) or '-' not in str(str_turno): return None, None
        parts = str(str_turno).split('-')
        return datetime.strptime(parts[0].strip(), "%H:%M").time(), datetime.strptime(parts[1].strip(), "%H:%M").time()
    except: return None, None

def calcular_minutos_exactos(row):
    """Calcula minutos de diferencia (retraso)."""
    try:
        turno = row.get('Turno')
        hora_real = row.get('Hora Entrada')
        fecha = row.get('Fecha Entrada')
        
        # Si no hay hora, es 0 para el cálculo de minutos (se maneja como ausencia en otro lado)
        if pd.isna(hora_real) or str(hora_real).strip() in ['-', '', 'nan']:
            return 0
        
        t_ini, t_fin = parsear_turno(turno)
        if t_ini is None: return 0

        # Parsear fecha
        fecha_dt = pd.to_datetime(fecha, dayfirst=True).date()
        
        # Parsear hora real (Manejo de formatos HH:MM:SS y HH:MM)
        h_str = str(hora_real).strip()
        try: h_real = datetime.strptime(h_str, "%H:%M:%S").time()
        except: 
            try: h_real = datetime.strptime(h_str, "%H:%M").time()
            except: return 0 # Formato desconocido
        
        # Lógica Nocturna / Cruce de Día
        dt_teorico = datetime.combine(fecha_dt, t_ini)
        
        # Si el turno empieza tarde (ej 22:00) y la marca es temprano (ej 00:15), es día siguiente
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
    """(4) Genera texto legible: Recinto | Hora | Estado"""
    hora_real = row.get('Hora Entrada')
    recinto = row.get('Dentro del Recinto (Entrada)', 'N/D')
    
    # Validación estricta de falta de marca
    if pd.isna(hora_real) or str(hora_real).strip() in ['-', '', 'nan']:
        return "Sin Marcaje [No registró entrada]"
    
    minutos_retraso = calcular_minutos_exactos(row)
    h_str = str(hora_real).strip()[:5]
    
    if minutos_retraso > tolerancia:
        estado = f"{minutos_retraso} min retraso"
    else:
        estado = "OK"
        
    return f"Recinto: {recinto} | Hora: {h_str} [{estado}]"

def formatear_detalle_salida(row):
    """(5) Genera texto legible salida"""
    hora_salida = row.get('Hora Salida')
    recinto = row.get('Dentro del Recinto (Salida)', 'N/D')
    
    if pd.isna(hora_salida) or str(hora_salida).strip() in ['-', '', 'nan']:
        return f"Recinto: {recinto} | Hora: Sin Marca"
    
    return f"Recinto: {recinto} | Hora: {str(hora_salida)[:5]}"

def determinar_tipo_preliminar(row, tolerancia):
    """(6) Tipo Preliminar"""
    hora_real = row.get('Hora Entrada')
    if pd.isna(hora_real) or str(hora_real).strip() in ['-', '', 'nan']:
        return "Inasistencia"
    
    if calcular_minutos_exactos(row) > tolerancia:
        return "Incidencia"
    
    return "OK"

# --- INTERFAZ PRINCIPAL ---

st.title("✈️ Plataforma de Gestión Operativa")
st.markdown("---")

with st.sidebar:
    st.header("1. Inputs")
    # Aclaración visual para el usuario
    st.info("Nota: El reporte de asistencia debe ser el formato original de BUK (inicia en fila 3).")
    
    file_asist = st.file_uploader("Reporte Asistencia (XLS/CSV)", type=["xls", "xlsx", "csv"])
    file_ina = st.file_uploader("Reporte Inasistencias (XLS/CSV)", type=["xls", "xlsx", "csv"])
    
    st.divider()
    st.header("2. Reglas")
    tolerancia = st.number_input("Tolerancia (min)", value=15, min_value=0)

if file_asist and file_ina:
    # ---------------------------------------------------------
    # 1. CARGA DE DATOS (CORRECCIÓN APLICADA AQUÍ)
    # ---------------------------------------------------------
    
    # Asistencia: Saltamos 2 filas porque los encabezados están en la fila 3 (índice 2)
    df_a = cargar_datos(file_asist, saltar_filas=2)
    
    # Inasistencias: Formato estándar (fila 1), no saltamos filas
    df_i = cargar_datos(file_ina, saltar_filas=0)
    
    # Validación de seguridad: Verificar si cargaron bien
    if df_a is not None and df_i is not None:
        
        # Normalizar nombres de columnas (quitar espacios extra)
        df_a.columns = df_a.columns.str.strip()
        df_i.columns = df_i.columns.str.strip()
        
        # DEBUG: Si el usuario quiere ver si se leyeron bien las columnas
        # st.write("Columnas detectadas en Asistencia:", df_a.columns.tolist())

        # ---------------------------------------------------------
        # 2. FILTRO ESPECIALIDAD
        # ---------------------------------------------------------
        all_specs = sorted(list(set(df_a['Especialidad'].dropna()) | set(df_i['Especialidad'].dropna())))
        selected_specs = st.sidebar.multiselect("Filtrar Especialidad", all_specs, default=all_specs)
        
        df_a = df_a[df_a['Especialidad'].isin(selected_specs)].copy()
        df_i = df_i[df_i['Especialidad'].isin(selected_specs)].copy()

        # ---------------------------------------------------------
        # 3. PREPARACIÓN Y UNIÓN
        # ---------------------------------------------------------
        df_a['Origen'] = 'Asistencia'
        df_i['Origen'] = 'Inasistencia'
        
        # Mapeo de columnas de Inasistencia BUK
        df_i['Fecha Entrada'] = df_i.get('Día', '') 
        df_i['Hora Entrada'] = '-'
        df_i['Hora Salida'] = '-'
        
        # Columnas maestras necesarias
        cols_comunes = ['RUT', 'Nombre', 'Primer Apellido', 'Especialidad', 'Turno', 
                        'Fecha Entrada', 'Hora Entrada', 'Hora Salida', 
                        'Dentro del Recinto (Entrada)', 'Dentro del Recinto (Salida)', 'Origen']
        
        # Rellenar columnas faltantes con None
        for c in cols_comunes:
            if c not in df_a.columns: df_a[c] = None
            if c not in df_i.columns: df_i[c] = None
                
        df_master = pd.concat([df_a[cols_comunes], df_i[cols_comunes]], ignore_index=True)

        # ---------------------------------------------------------
        # 4. CÁLCULO DE ATRIBUTOS (LOS 9 PUNTOS)
        # ---------------------------------------------------------
        
        # (1) Colaborador
        df_master['Colaborador'] = df_master['Nombre'].astype(str) + " " + df_master['Primer Apellido'].astype(str)
        
        # (9) Minutos Incidencia (Pre-cálculo antes de formatear textos)
        df_master['Minutos Incidencia'] = df_master.apply(calcular_minutos_exactos, axis=1)

        # (4) Detalle Entrada
        df_master['Detalle Entrada'] = df_master.apply(lambda x: formatear_detalle_entrada(x, tolerancia), axis=1)
        
        # (5) Detalle Salida
        df_master['Detalle Salida'] = df_master.apply(formatear_detalle_salida, axis=1)
        
        # (6) Tipo
        df_master['Tipo'] = df_master.apply(lambda x: determinar_tipo_preliminar(x, tolerancia), axis=1)
        
        # (7) Pre-Clasificación
        def pre_clasificar(row):
            tipo = row['Tipo']
            detalle = row['Detalle Entrada']
            if tipo == 'Inasistencia': return 'Ausencia'
            if 'retraso' in detalle: return 'Retraso'
            if 'Anticipada' in detalle: return 'Salida Anticipada' # Si lo implementamos en detalle
            return 'OK'

        df_master['Clasificación'] = df_master.apply(pre_clasificar, axis=1)
        
        # (8) Justificación Default
        df_master['Justificación'] = 'Injustificado'

        # ---------------------------------------------------------
        # 5. VISUALIZACIÓN (TABLA EDITABLE)
        # ---------------------------------------------------------
        
        # Seleccionar columnas finales
        cols_finales = [
            'Colaborador', 'Fecha Entrada', 'Turno', 'Detalle Entrada', 'Detalle Salida',
            'Tipo', 'Clasificación', 'Justificación', 'Minutos Incidencia'
        ]
        
        df_visual = df_master[cols_finales].copy()
        
        st.subheader("📝 Detalle de Incidencias")
        st.info("Validación: Se han omitido las primeras 2 filas del reporte de Asistencia para leer correctamente los encabezados.")

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
                "Minutos Incidencia": st.column_config.NumberColumn(
                    "9. Min. Real",
                    min_value=0, step=1, required=True
                )
            },
            hide_index=True,
            num_rows="fixed",
            use_container_width=True,
            height=600
        )

        # ---------------------------------------------------------
        # 6. REPORTING KPI
        # ---------------------------------------------------------
        st.divider()
        st.subheader("📊 Reporte de KPIs")

        # Excluir 'No Procede'
        df_calc = edited_df[edited_df['Clasificación'] != 'No Procede'].copy()
        
        if len(df_calc) > 0:
            total = len(df_calc)
            
            # Minutos a descontar (Solo Injustificados)
            minutos_total = df_calc[df_calc['Justificación'] == 'Injustificado']['Minutos Incidencia'].sum()
            
            # Cantidad de Problemas (Para % Cumplimiento)
            # Problema = (No es OK) Y (Es Injustificado)
            problemas = df_calc[
                (df_calc['Clasificación'] != 'OK') & 
                (df_calc['Justificación'] == 'Injustificado')
            ]
            
            cumplimiento = ((total - len(problemas)) / total) * 100
            
            k1, k2, k3 = st.columns(3)
            k1.metric("Turnos Procesados", total)
            k2.metric("Minutos Descuento", int(minutos_total), help="Suma de Minutos Reales Injustificados")
            k3.metric("Cumplimiento", f"{cumplimiento:.1f}%")
            
            # Exportar Excel
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                edited_df.to_excel(writer, index=False, sheet_name='Detalle Validado')
                
                # Hoja extra con resumen para Nómina
                resumen_nomina = df_calc[df_calc['Justificación'] == 'Injustificado'][['Colaborador', 'Fecha Entrada', 'Minutos Incidencia', 'Clasificación']]
                resumen_nomina.to_excel(writer, index=False, sheet_name='Resumen Nómina')
            
            st.download_button("💾 Descargar Excel", buffer.getvalue(), "Reporte_Final.xlsx")
            
    else:
        st.warning("No se pudieron leer los archivos. Verifica que el archivo de Asistencia sea el formato BUK correcto.")

else:
    st.info("Carga 'Asistencia' e 'Inasistencia' en el menú lateral.")
