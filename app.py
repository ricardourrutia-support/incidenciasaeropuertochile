import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import io
import plotly.express as px

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(
    page_title="Gestión de Ausentismo e Incidencias",
    page_icon="✈️",
    layout="wide"
)

# --- ESTILOS CSS PERSONALIZADOS (Manteniendo tu estilo visual) ---
st.markdown("""
    <style>
    .main {
        background-color: #f8f9fa;
    }
    .stMetric {
        background-color: #ffffff;
        padding: 15px;
        border-radius: 5px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    </style>
    """, unsafe_allow_html=True)

# --- FUNCIONES DE LÓGICA DE NEGOCIO (EL NUEVO MOTOR) ---

def cargar_archivo(uploaded_file):
    """Detecta si es CSV o Excel y lo carga en DataFrame."""
    if uploaded_file is None:
        return None
    try:
        if uploaded_file.name.endswith('.csv'):
            # Intentamos detectar separador automáticamente
            try:
                return pd.read_csv(uploaded_file)
            except:
                return pd.read_csv(uploaded_file, sep=';')
        else:
            return pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"Error cargando {uploaded_file.name}: {e}")
        return None

def parsear_turno(str_turno):
    """
    Extrae hora inicio y fin de strings como '20:00-07:00'.
    Retorna objetos time.
    """
    try:
        if pd.isna(str_turno) or '-' not in str(str_turno):
            return None, None
        parts = str(str_turno).split('-')
        t_ini = datetime.strptime(parts[0].strip(), "%H:%M").time()
        t_fin = datetime.strptime(parts[1].strip(), "%H:%M").time()
        return t_ini, t_fin
    except:
        return None, None

def calcular_estado_asistencia(row, tolerancia_min):
    """
    Compara hora real vs turno planificado.
    Maneja lógica de turno nocturno (cruce de medianoche).
    """
    # 1. Datos base
    fecha_str = row.get('Fecha Entrada')
    hora_real_str = row.get('Hora Entrada')
    str_turno = row.get('Turno')
    
    # 2. Validaciones básicas
    if pd.isna(hora_real_str):
        return "Sin Marcaje (Ausencia o Error)"
    if pd.isna(str_turno):
        return "Sin Turno Asignado"

    try:
        # 3. Parsear Fecha Base y Hora Real
        fecha_base = pd.to_datetime(fecha_str, dayfirst=True).date()
        
        # Limpieza de hora real (a veces viene como HH:MM:SS)
        hora_real_clean = str(hora_real_str).strip()
        try:
            hora_real_time = datetime.strptime(hora_real_clean, "%H:%M:%S").time()
        except:
            hora_real_time = datetime.strptime(hora_real_clean, "%H:%M").time()
            
        dt_real = datetime.combine(fecha_base, hora_real_time)

        # 4. Parsear Turno Teórico
        t_ini, t_fin = parsear_turno(str_turno)
        if t_ini is None:
            return "Error Formato Turno"

        dt_teorico_ini = datetime.combine(fecha_base, t_ini)
        
        # AJUSTE NOCTURNO:
        # Si la persona tiene turno 22:00, pero marca a las 00:15 (ya es día siguiente en la realidad)
        # O si marca a las 21:50 (mismo día).
        # La lógica simple es: Comparamos cercanía.
        
        # Calculamos diferencia en minutos
        diff_minutes = (dt_real - dt_teorico_ini).total_seconds() / 60
        
        # Corrección para cruce de día en marcaje real vs teórico
        # Si diff es ej: -1400 min (marcó dia siguiente?) o +1400 min.
        # Asumiremos que BUK entrega la fecha correcta del marcaje. 
        
        if diff_minutes > tolerancia_min:
            return f"Retraso ({int(diff_minutes)} min)"
        elif diff_minutes < -60:
             return "Llegada Anticipada" # O error de fecha
        else:
            return "Asistencia Correcta"
            
    except Exception as e:
        return "Error Cálculo"

# --- INTERFAZ DE USUARIO ---

st.title("✈️ Plataforma de Gestión de Ausentismo")
st.markdown("Consolidación de Asistencias e Inasistencias (BUK) para reporte de nómina.")
st.markdown("---")

# 1. SIDEBAR: CARGA Y CONFIGURACIÓN
with st.sidebar:
    st.header("1. Carga de Datos")
    f_asist = st.file_uploader("Cargar Asistencia (.csv/.xls)", type=["csv", "xls", "xlsx"])
    f_inasist = st.file_uploader("Cargar Inasistencias (.csv/.xls)", type=["csv", "xls", "xlsx"])
    
    st.markdown("---")
    st.header("2. Filtros y Reglas")
    
    # Selector de fechas (Opcional, filtra sobre los datos cargados)
    # Nota: Se activará cuando haya datos
    
    tolerancia = st.number_input("Tolerancia de Retraso (minutos)", value=15, min_value=0, step=1)

# 2. PROCESAMIENTO PRINCIPAL
if f_asist and f_inasist:
    
    # --- ETL ---
    df_asist = cargar_archivo(f_asist)
    df_inasist = cargar_archivo(f_inasist)
    
    # Normalización de columnas (Eliminar espacios en nombres)
    df_asist.columns = df_asist.columns.str.strip()
    df_inasist.columns = df_inasist.columns.str.strip()

    # --- FILTRO DE ESPECIALIDAD ---
    # Unimos todas las especialidades para llenar el selector
    specs_asist = df_asist['Especialidad'].dropna().unique().tolist()
    specs_ina = df_inasist['Especialidad'].dropna().unique().tolist()
    all_specs = sorted(list(set(specs_asist + specs_ina)))
    
    with st.sidebar:
        selected_specs = st.multiselect(
            "Filtrar por Especialidad", 
            options=all_specs,
            default=all_specs
        )
    
    # Filtrar Dataframes
    df_asist = df_asist[df_asist['Especialidad'].isin(selected_specs)].copy()
    df_inasist = df_inasist[df_inasist['Especialidad'].isin(selected_specs)].copy()
    
    if df_asist.empty and df_inasist.empty:
        st.warning("No hay datos para las especialidades seleccionadas.")
    else:
        # --- LÓGICA DE NEGOCIO ---
        
        # 1. Procesar Asistencias (Detectar Retrasos)
        # Usamos la función que creamos arriba
        df_asist['Estado'] = df_asist.apply(lambda row: calcular_estado_asistencia(row, tolerancia), axis=1)
        
        # 2. Procesar Inasistencias
        # Estandarizamos para que coincida con la tabla maestra
        df_inasist['Estado'] = "Inasistencia: " + df_inasist['Motivo'].astype(str)
        # Llenamos columnas faltantes para el merge
        df_inasist['Hora Entrada'] = "-"
        df_inasist['Fecha Entrada'] = df_inasist.get('Día', '') # A veces BUK llama a la fecha 'Día' en inasistencias
        
        # 3. Unificar (Master Table)
        # Definimos las columnas clave que queremos ver
        cols_clave = ['RUT', 'Nombre', 'Primer Apellido', 'Especialidad', 'Turno', 'Fecha Entrada', 'Hora Entrada', 'Estado']
        
        # Asegurar que existan en ambos (rellenar si falta alguna)
        for col in cols_clave:
            if col not in df_asist.columns: df_asist[col] = ""
            if col not in df_inasist.columns: df_inasist[col] = ""
            
        df_master = pd.concat([
            df_asist[cols_clave],
            df_inasist[cols_clave]
        ], ignore_index=True)
        
        # Añadir columnas para interacción del Supervisor
        df_master['Justificación'] = "Pendiente"
        df_master['Es_Justificado'] = False # Checkbox
        
        # --- DASHBOARD VISUAL ---
        
        st.subheader("📊 Resumen General")
        
        # Métricas simples preliminares
        total_records = len(df_master)
        total_retrasos = len(df_master[df_master['Estado'].str.contains("Retraso", na=False)])
        total_ausencias = len(df_master[df_master['Estado'].str.contains("Inasistencia", na=False)])
        
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Registros", total_records)
        col2.metric("Retrasos Detectados", total_retrasos, delta_color="inverse")
        col3.metric("Inasistencias", total_ausencias, delta_color="inverse")
        
        # --- TABLA INTERACTIVA (DATA EDITOR) ---
        st.markdown("### 📝 Gestión de Incidencias")
        st.info("Utiliza esta tabla para justificar inasistencias o retrasos. Los cambios recalculan el cumplimiento.")
        
        edited_df = st.data_editor(
            df_master,
            column_config={
                "Es_Justificado": st.column_config.CheckboxColumn(
                    "¿Justificado?",
                    help="Marcar si esta incidencia no debe contar para el descuento.",
                    default=False,
                ),
                "Justificación": st.column_config.SelectboxColumn(
                    "Motivo Justificación",
                    options=[
                        "Pendiente",
                        "Licencia Médica", 
                        "Permiso Legal", 
                        "Falla Transporte", 
                        "Error de Marcaje", 
                        "Cambio Turno Autorizado",
                        "Injustificado"
                    ],
                    required=True,
                ),
                "Estado": st.column_config.TextColumn(
                    "Estado Sistema",
                    width="medium",
                    disabled=True
                )
            },
            disabled=["RUT", "Nombre", "Turno", "Hora Entrada"], # Bloquear edición de datos originales
            hide_index=True,
            use_container_width=True,
            height=600
        )
        
        # --- CÁLCULO FINAL DE CUMPLIMIENTO ---
        # El cumplimiento se calcula sobre el dataframe EDITADO
        
        # Definimos "Incidencia Real" como: Estado != Correcto Y No Justificado
        def es_incidencia_final(row):
            estado = str(row['Estado'])
            justificado = row['Es_Justificado']
            
            if justificado:
                return False # Si está justificado, no resta cumplimiento
            
            # Si es Inasistencia o Retraso o Sin Marcaje -> Es Incidencia
            if "Inasistencia" in estado or "Retraso" in estado or "Sin Marcaje" in estado:
                return True
            return False

        incidencias_finales = edited_df.apply(es_incidencia_final, axis=1).sum()
        cumplimiento_pct = ((total_records - incidencias_finales) / total_records) * 100 if total_records > 0 else 0
        
        st.divider()
        st.subheader("📈 Cumplimiento Final")
        
        c_kpi1, c_kpi2 = st.columns([1, 3])
        c_kpi1.metric("Incidencias Finales (Injustificadas)", int(incidencias_finales))
        c_kpi1.metric("% Cumplimiento Operativo", f"{cumplimiento_pct:.1f}%")
        
        # Gráfico simple de barras por Estado
        conteo_estados = edited_df['Estado'].value_counts().reset_index()
        conteo_estados.columns = ['Tipo Incidencia', 'Cantidad']
        fig = px.bar(conteo_estados, x='Tipo Incidencia', y='Cantidad', title="Distribución de Incidencias")
        c_kpi2.plotly_chart(fig, use_container_width=True)

        # --- EXPORTACIÓN ---
        st.markdown("### 📥 Descargar Reporte Validado")
        
        # Convertir a Excel para descarga (Mejor formato para nómina)
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            edited_df.to_excel(writer, index=False, sheet_name='Reporte Consolidado')
            # Auto-adjust columns width (opcional, visual)
        
        st.download_button(
            label="Descargar Excel (.xlsx)",
            data=buffer.getvalue(),
            file_name=f"Reporte_Ausentismo_Validado_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.ms-excel"
        )

else:
    # PANTALLA DE INICIO (Cuando no hay archivos)
    st.info("👈 Por favor, carga los reportes de **Asistencia** e **Inasistencia** en el menú lateral para comenzar.")
    st.markdown("""
    ### Instrucciones:
    1. Descarga los reportes desde BUK en formato Excel o CSV.
    2. Súbelos en la barra lateral izquierda.
    3. Ajusta la **Tolerancia** si es necesario.
    4. Usa la tabla central para **Justificar** incidencias.
    5. Descarga el reporte final validado.
    """)
