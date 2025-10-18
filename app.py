import streamlit as st
import pandas as pd
import re
from io import BytesIO

# Configuración de la página
st.set_page_config(page_title="Generador Planner & Gabinete", page_icon="📊", layout="wide")

# Título de la aplicación
st.title("📊 Generador de Archivos Planner & Gabinete")
st.markdown("---")

# Diccionarios de regiones y responsables
regiones = {
    'R01': 'Región de Tarapacá', 'R02': 'Región de Antofagasta', 'R03': 'Región de Atacama',
    'R04': 'Región de Coquimbo', 'R05': 'Región de Valparaíso', 
    'R06': 'Región del Libertador Bernardo O\'Higgins', 'R07': 'Región del Maule',
    'R08': 'Región del Biobío', 'R09': 'Región de La Araucanía', 'R10': 'Región de Los Lagos',
    'R11': 'Región Aysén', 'R12': 'Región de Magallanes', 'R13': 'Región Metropolitana',
    'R14': 'Región de Los Ríos', 'R15': 'Región de Arica y Parinacota', 
    'R16': 'Región de Ñuble', 'R00': 'Nacional'
}

responsables = {
    'ODLF': 'Osvaldo de la Fuente Castro', 'ANG': 'Alvaro Núñez Gómez de Jiménez',
    'AGL': 'Antonia Garrido', 'CBC': 'Catalina Berrios', 
    'CLV': 'Constanza Lavanderos Vergara', 'CRT': 'Constanza Reyes',
    'KOR': 'Karla Orrego Romero', 'LMS': 'Lukas Moenne Saito',
    'MMA': 'Maria Isabel Mallea Alvarez', 'RHA': 'Raul Herrera Araya',
    'JAA': 'Jorge Alviña Aguayo', 'FBB': 'Francisca Barrios Benavente'
}

# Función para limpiar etiquetas
def limpiar_etiquetas(cadena):
    if pd.isna(cadena):
        return cadena
    etiquetas = [e.strip() for e in cadena.split(';')]
    etiquetas_filtradas = [e for e in etiquetas if e.lower() not in ['acto normal', 'acto complejo', 'acto simple']]
    return '; '.join(etiquetas_filtradas) if etiquetas_filtradas else None

# Función para procesar los datos
def procesar_datos(df):
    # Procesamiento igual al original
    df['codigo'] = df['Nombre de la tarea'].str.extract(r'(R\d{2})')
    df['Región'] = df['codigo'].map(regiones)
    df.drop(columns='codigo', inplace=True)

    df['codigo_resp'] = df['Nombre de la tarea'].str.extract(r'^(.*?)\*\*')
    df['Responsable'] = df['codigo_resp'].map(responsables)
    df.drop(columns='codigo_resp', inplace=True)

    # Limpiar nombre de la tarea
    pattern = r'\w+\*\*\w+\*\*R([0-1][1-9]|[0-1][0-6])\*\* '
    df['Nombre de la tarea'] = df['Nombre de la tarea'].str.replace(pattern, '', regex=True)

    # Filtrar tareas no completadas
    df = df[df['Progreso'] != 'Completado']

    # Seleccionar columnas relevantes
    df = df[['Responsable','Nombre de la tarea','Región','Etiquetas','Nombre del depósito','Progreso','Priority','Asignado a']]

    # Aplicar limpieza de etiquetas
    df['Etiquetas'] = df['Etiquetas'].apply(limpiar_etiquetas)
    
    return df

# Función para crear el Excel en memoria
def crear_excel_en_memoria(df, fecha_semana):
    output = BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Hoja 1: "planner" - Datos procesados completos
        df.to_excel(writer, sheet_name='planner', index=False)
        
        # Hoja 2: "gabinete" - Datos filtrados
        gabinete_df = df.copy()
        
        # Aplicar filtros para gabinete
        depositos_validos = ['Archivo', 'Caducidades', 'Elusiones', 'Medidas', 'Seguimiento Ambiental']
        gabinete_df = gabinete_df[
            (gabinete_df['Nombre del depósito'].isin(depositos_validos)) &
            (gabinete_df['Priority'] != 'Baja') &
            (gabinete_df['Responsable'].str.contains('Bruno', na=False))
        ]
        
        # Si hay datos después del filtro, crear la hoja gabinete
        if len(gabinete_df) > 0:
            gabinete_final = pd.DataFrame({
                'Semana': [fecha_semana] * len(gabinete_df),
                'Área (DSC/FIS)': 'FIS',
                'Responsable': gabinete_df['Responsable'].values,
                'UF/Proyecto': gabinete_df['Nombre de la tarea'].values,
                'Región': gabinete_df['Región'].values,
                'Producto': gabinete_df['Etiquetas'].values,
                'IGA (RCA/Ruido/PPDA/Lumínica/RILes)': gabinete_df['Nombre del depósito'].values
            })
        else:
            # Si no hay datos, crear DataFrame vacío con las columnas
            gabinete_final = pd.DataFrame(columns=[
                'Semana', 'Área (DSC/FIS)', 'Responsable', 'UF/Proyecto', 
                'Región', 'Producto', 'IGA (RCA/Ruido/PPDA/Lumínica/RILes)'
            ])
        
        # Guardar hoja gabinete
        gabinete_final.to_excel(writer, sheet_name='gabinete', index=False)
    
    output.seek(0)
    return output

# Interfaz de usuario
st.header("📤 Cargar Archivo Excel")

uploaded_file = st.file_uploader("Selecciona el archivo Excel de tareas", type=['xlsx', 'xls'])

if uploaded_file is not None:
    try:
        # Leer el archivo
        df_original = pd.read_excel(uploaded_file)
        
        st.success(f"✅ Archivo cargado correctamente: {uploaded_file.name}")
        st.info(f"📊 Registros encontrados: {len(df_original)}")
        
        # Mostrar vista previa
        with st.expander("🔍 Vista previa de los datos originales"):
            st.dataframe(df_original.head())
        
        # Input para la fecha
        st.header("📅 Configuración de Gabinete")
        fecha_semana = st.text_input("Ingresa la fecha para la columna Semana", value="14/10/2024")
        
        # Botón para procesar
        if st.button("🚀 Generar Archivo Excel", type="primary"):
            with st.spinner("Procesando datos..."):
                # Procesar datos
                df_procesado = procesar_datos(df_original)
                
                # Crear Excel en memoria
                excel_output = crear_excel_en_memoria(df_procesado, fecha_semana)
                
                # Mostrar estadísticas
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Registros Planner", len(df_procesado))
                with col2:
                    gabinete_count = len(df_procesado[
                        (df_procesado['Nombre del depósito'].isin(['Archivo', 'Caducidades', 'Elusiones', 'Medidas', 'Seguimiento Ambiental'])) &
                        (df_procesado['Priority'] != 'Baja') &
                        (df_procesado['Responsable'].str.contains('Bruno', na=False))
                    ])
                    st.metric("Registros Gabinete", gabinete_count)
                with col3:
                    st.metric("Fecha Semana", fecha_semana)
                
                # Mostrar vistas previas
                col1, col2 = st.columns(2)
                
                with col1:
                    st.subheader("Hoja Planner")
                    st.dataframe(df_procesado.head())
                
                with col2:
                    st.subheader("Hoja Gabinete")
                    gabinete_filtrado = df_procesado[
                        (df_procesado['Nombre del depósito'].isin(['Archivo', 'Caducidades', 'Elusiones', 'Medidas', 'Seguimiento Ambiental'])) &
                        (df_procesado['Priority'] != 'Baja') &
                        (df_procesado['Responsable'].str.contains('Bruno', na=False))
                    ]
                    if len(gabinete_filtrado) > 0:
                        gabinete_preview = pd.DataFrame({
                            'Semana': [fecha_semana] * len(gabinete_filtrado.head()),
                            'Área': 'FIS',
                            'Responsable': gabinete_filtrado['Responsable'].head(),
                            'UF/Proyecto': gabinete_filtrado['Nombre de la tarea'].head(),
                            'Región': gabinete_filtrado['Región'].head(),
                            'Producto': gabinete_filtrado['Etiquetas'].head(),
                            'IGA': gabinete_filtrado['Nombre del depósito'].head()
                        })
                        st.dataframe(gabinete_preview)
                    else:
                        st.info("No hay registros que cumplan los criterios de filtro para Gabinete")
                
                # Botón de descarga
                st.download_button(
                    label="📥 Descargar Archivo Excel",
                    data=excel_output,
                    file_name=f"planner_gabinete_{fecha_semana.replace('/', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
                
    except Exception as e:
        st.error(f"❌ Error al procesar el archivo: {e}")
else:
    st.info("👆 Por favor, carga un archivo Excel para comenzar")

# Información adicional
with st.expander("ℹ️ Información sobre el procesamiento"):
    st.markdown("""
    **Procesamiento aplicado:**
    - Extracción de región desde código R## en el nombre de tarea
    - Identificación de responsable desde iniciales
    - Limpieza de patrones en nombres de tareas
    - Filtrado de tareas no completadas
    - Limpieza de etiquetas (remoción de 'acto normal', 'acto complejo', 'acto simple')
    
    **Filtros para Gabinete:**
    - Nombre del depósito: Archivo, Caducidades, Elusiones, Medidas, Seguimiento Ambiental
    - Priority: No 'Baja'
    - Responsable: Contiene 'Bruno'
    """)
