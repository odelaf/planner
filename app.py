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

# Función mejorada para crear el Excel en memoria con depuración
def crear_excel_en_memoria(df, fecha_semana):
    output = BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Hoja 1: "planner" - Datos procesados completos
        df.to_excel(writer, sheet_name='planner', index=False)
        
        # Hoja 2: "gabinete" - Datos filtrados
        gabinete_df = df.copy()
        
        # DEPURACIÓN: Mostrar valores únicos para diagnóstico
        st.subheader("🔍 Depuración - Valores Únicos")
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.write("**Nombre del depósito:**")
            st.write(gabinete_df['Nombre del depósito'].unique())
        
        with col2:
            st.write("**Priority:**")
            st.write(gabinete_df['Priority'].unique())
        
        with col3:
            st.write("**Responsable:**")
            st.write(gabinete_df['Responsable'].unique())
        
        # FILTROS MEJORADOS
        # 1. Filtro de depósitos (case insensitive y contiene)
        depositos_validos = ['Archivo', 'Caducidades', 'Elusiones', 'Medidas', 'Seguimiento Ambiental']
        
        # Opción A: Filtro exacto (como estaba)
        filtro_depositos_exacto = gabinete_df['Nombre del depósito'].isin(depositos_validos)
        
        # Opción B: Filtro por contiene (por si hay variaciones)
        filtro_depositos_contiene = gabinete_df['Nombre del depósito'].apply(
            lambda x: any(deposito in str(x) for deposito in depositos_validos) if pd.notna(x) else False
        )
        
        # Usar el filtro que funcione mejor
        filtro_depositos = filtro_depositos_exacto | filtro_depositos_contiene
        
        # 2. Filtro de priority (excluir 'Baja')
        filtro_priority = gabinete_df['Priority'] != 'Baja'
        
        # 3. Filtro de responsable (case insensitive, contiene 'Bruno')
        filtro_responsable = gabinete_df['Responsable'].str.contains('Bruno', case=False, na=False)
        
        # Aplicar filtros combinados
        gabinete_filtrado = gabinete_df[filtro_depositos & filtro_priority & filtro_responsable]
        
        # Mostrar resultados de cada filtro
        st.subheader("📊 Resultados de Filtros")
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric("Total registros", len(gabinete_df))
        with col2:
            st.metric("Filtro depósitos", filtro_depositos.sum())
        with col3:
            st.metric("Filtro priority", filtro_priority.sum())
        with col4:
            st.metric("Filtro responsable", filtro_responsable.sum())
        
        st.metric("✅ Registros después de filtros", len(gabinete_filtrado))
        
        # Si hay datos después del filtro, crear la hoja gabinete
        if len(gabinete_filtrado) > 0:
            gabinete_final = pd.DataFrame({
                'Semana': [fecha_semana] * len(gabinete_filtrado),
                'Área (DSC/FIS)': 'FIS',
                'Responsable': gabinete_filtrado['Responsable'].values,
                'UF/Proyecto': gabinete_filtrado['Nombre de la tarea'].values,
                'Región': gabinete_filtrado['Región'].values,
                'Producto': gabinete_filtrado['Etiquetas'].values,
                'IGA (RCA/Ruido/PPDA/Lumínica/RILes)': gabinete_filtrado['Nombre del depósito'].values
            })
            
            # Mostrar preview del gabinete
            st.subheader("👀 Vista previa - Gabinete Filtrado")
            st.dataframe(gabinete_final)
            
        else:
            # Si no hay datos, crear DataFrame vacío con las columnas
            gabinete_final = pd.DataFrame(columns=[
                'Semana', 'Área (DSC/FIS)', 'Responsable', 'UF/Proyecto', 
                'Región', 'Producto', 'IGA (RCA/Ruido/PPDA/Lumínica/RILes)'
            ])
            st.warning("⚠️ No se encontraron registros que cumplan todos los criterios de filtro")
            
            # Mostrar posibles problemas
            st.info("💡 **Posibles causas:**")
            st.write("- No hay registros con 'Bruno' en Responsable")
            st.write("- Los depósitos no coinciden exactamente con los valores esperados")
            st.write("- Todos los registros tienen Priority 'Baja'")
            st.write("- Valores nulos en las columnas de filtro")
        
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
                
                # Mostrar estadísticas iniciales
                st.subheader("📈 Estadísticas Iniciales")
                col1, col2 = st.columns(2)
                with col1:
                    st.metric("Registros Planner", len(df_procesado))
                with col2:
                    st.metric("Tareas No Completadas", len(df_procesado))
                
                # Crear Excel en memoria con depuración
                excel_output = crear_excel_en_memoria(df_procesado, fecha_semana)
                
                # Botón de descarga
                st.download_button(
                    label="📥 Descargar Archivo Excel",
                    data=excel_output,
                    file_name=f"planner_gabinete_{fecha_semana.replace('/', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
                
    except Exception as e:
        st.error(f"❌ Error al procesar el archivo: {str(e)}")
        st.info("💡 Asegúrate de que el archivo tenga la estructura correcta")
else:
    st.info("👆 Por favor, carga un archivo Excel para comenzar")
