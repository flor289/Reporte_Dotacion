import streamlit as st
import pandas as pd
import plotly.express as px
from fpdf import FPDF
from datetime import datetime
import io

# --- CONFIGURACIÓN Y ESTILOS ---
st.set_page_config(page_title="Dashboard RRHH - Florencia Flores", layout="wide")

st.markdown("""
    <style>
    .main { background-color: #f0f2f6; }
    h1, h2, h3 { color: #003366; }
    div.stDownloadButton > button { background-color: #28a745; color: white; }
    </style>
    """, unsafe_allow_html=True)

# --- FUNCIONES DE PROCESAMIENTO ---

def procesar_todo(archivo):
    # 1. Carga de pestañas
    df_base = pd.read_excel(archivo, sheet_name='BaseQuery')
    df_activos_viejos = pd.read_excel(archivo, sheet_name='Activos')
    try:
        df_co_manual = pd.read_excel(archivo, sheet_name='CO')
    except:
        df_co_manual = pd.DataFrame(columns=['Nº pers.', 'Apellido', 'Nombre de pila', 'Línea', 'Categoría', 'Desde', 'Motivo'])

    # Estandarizar legajos
    for df in [df_base, df_activos_viejos, df_co_manual]:
        if 'Nº pers.' in df.columns:
            df['Nº pers.'] = df['Nº pers.'].astype(str).str.strip()

    # 2. Identificar Salidas por Comparación
    legajos_viejos = set(df_activos_viejos['Nº pers.'])
    legajos_nuevos = set(df_base[df_base['Status ocupación'] == 'Activo']['Nº pers.'])
    
    # IDs de personas que salieron (estaban antes y ahora no están o están de baja)
    ids_salidas = legajos_viejos - legajos_nuevos

    # 3. Clasificar: ¿Es Baja de Sistema o Cambio Organizativo (CO)?
    # Bajas: Están en BaseQuery pero con status "Dado de baja"
    df_bajas_sis = df_base[(df_base['Nº pers.'].isin(ids_salidas)) & (df_base['Status ocupación'] == 'Dado de baja')].copy()
    df_bajas_sis['Tipo'] = 'Baja'

    # CO: Estaban en Activos pero desaparecieron totalmente de BaseQuery
    ids_en_base = set(df_base['Nº pers.'])
    ids_co = ids_salidas - ids_en_base
    
    # Cruzamos los IDs de CO con tu pestaña manual para recuperar los nombres/motivos
    df_co_detectados = df_co_manual[df_co_manual['Nº pers.'].isin(ids_co)].copy()
    df_co_detectados['Tipo'] = 'Cambio Organizativo'
    if 'Motivo' in df_co_detectados.columns:
        df_co_detectados['Motivo de la medida'] = df_co_detectados['Motivo']

    # 4. Unificar Dataset de Salidas
    columnas_finales = ['Nº pers.', 'Apellido', 'Nombre de pila', 'Línea', 'Categoría', 'Desde', 'Motivo de la medida', 'Tipo']
    
    # Asegurar que existan columnas para el concat
    for col in columnas_finales:
        if col not in df_bajas_sis.columns: df_bajas_sis[col] = "Sin Datos"
        if col not in df_co_detectados.columns: df_co_detectados[col] = "Sin Datos"

    df_final = pd.concat([df_bajas_sis[columnas_finales], df_co_detectados[columnas_finales]], ignore_index=True)
    df_final['Desde'] = pd.to_datetime(df_final['Desde'], errors='coerce')
    
    return df_final, df_base

# --- INTERFAZ DE USUARIO ---

st.title("📊 Control de Gestión: Bajas y Cambios Organizativos")
st.write("Carga tu archivo con las pestañas `BaseQuery`, `Activos` y `CO` para comparar los estados.")

archivo_cargado = st.file_uploader("Subir archivo Excel (.xlsx)", type=['xlsx'])

if archivo_cargado:
    try:
        df_salidas, df_actual = procesar_todo(archivo_cargado)
        
        # --- CÁLCULO DE RANGO DINÁMICO ---
        if not df_salidas.empty and df_salidas['Desde'].notna().any():
            f_min = df_salidas['Desde'].min()
            f_max = df_salidas['Desde'].max()
            rango_str = f"{f_min.strftime('%d/%m/%Y')} al {f_max.strftime('%d/%m/%Y')}"
        else:
            rango_str = "No detectado"

        # TABS
        tab_evolucion, tab_detalle = st.tabs(["📉 Evolución Mensual", "👥 Detalle de Personas"])

        with tab_evolucion:
            st.header(f"Evolución Mensual de Bajas y Cambios Organizativos")
            st.info(f"Análisis basado en la comparación de archivos: **{rango_str}**")

            # KPIs
            c1, c2, c3 = st.columns(3)
            c1.metric("Total Salidas Detectadas", len(df_salidas))
            c2.metric("Bajas (Sistema)", len(df_salidas[df_salidas['Tipo'] == 'Baja']))
            c3.metric("Cambios Org. (CO)", len(df_salidas[df_salidas['Tipo'] == 'Cambio Organizativo']))

            # Tabla y Gráfico
            df_salidas['Mes-Año'] = df_salidas['Desde'].dt.strftime('%Y-%m')
            # Ordenar meses cronológicamente
            meses_ordenados = sorted(df_salidas['Mes-Año'].dropna().unique())
            
            resumen_mensual = pd.crosstab(df_salidas['Mes-Año'], df_salidas['Tipo'], margins=True, margins_name="Total")
            
            col_tabla, col_graf = st.columns([1, 2])
            with col_tabla:
                st.write("### 📅 Cuadro Mensual")
                st.dataframe(resumen_mensual, use_container_width=True)
            
            with col_graf:
                st.write("### 📈 Tendencia")
                fig = px.bar(df_salidas.dropna(subset=['Mes-Año']), x='Mes-Año', color='Tipo', 
                             barmode='group', category_orders={"Mes-Año": meses_ordenados},
                             color_discrete_map={'Baja': '#FF4B4B', 'Cambio Organizativo': '#1C83E1'})
                st.plotly_chart(fig, use_container_width=True)

        with tab_detalle:
            st.header("Listado Nominal de Salidas")
            st.write("Listado completo de personas detectadas como salida en la comparación.")
            
            # Buscador simple
            busqueda = st.text_input("Buscar por Apellido o Legajo:")
            df_mostrar = df_salidas.copy()
            if busqueda:
                df_mostrar = df_mostrar[df_mostrar.astype(str).apply(lambda x: x.str.contains(busqueda, case=False)).any(axis=1)]
            
            st.dataframe(df_mostrar.sort_values('Desde', ascending=False), hide_index=True, use_container_width=True)

    except Exception as e:
        st.error(f"Error al procesar el archivo: {e}")
        st.info("Asegúrate de que las pestañas se llamen exactamente: BaseQuery, Activos y CO.")
else:
    st.info("Sube un archivo para comenzar el análisis.")