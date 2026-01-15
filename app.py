import streamlit as st
import pandas as pd
import plotly.express as px
from fpdf import FPDF
from datetime import datetime, timedelta
import io

# --- CONFIGURACIÓN Y ESTILOS ---
COLOR_AZUL_INSTITUCIONAL = (4, 118, 208)
COLOR_TEXTO_TITULO = (0, 51, 102)
COLOR_TEXTO_CUERPO = (50, 50, 50)
COLOR_FONDO_CABECERA_TABLA = (70, 130, 180)
COLOR_GRIS_FONDO_FILA = (240, 242, 246)

class PDF(FPDF):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.page_width = self.w - 2 * self.l_margin
        self.report_title = ""

    def header(self):
        self.set_font("Arial", "B", 16)
        self.set_text_color(*COLOR_TEXTO_TITULO)
        self.cell(0, 10, self.report_title, 0, 1, "C")
        self.ln(5)

    def draw_table(self, title, df):
        if df.empty: return
        self.set_font("Arial", "B", 12)
        self.set_text_color(*COLOR_TEXTO_TITULO)
        self.cell(0, 10, title, ln=True)
        
        self.set_font("Arial", "B", 9)
        self.set_fill_color(*COLOR_FONDO_CABECERA_TABLA)
        self.set_text_color(255, 255, 255)
        
        # Calcular anchos
        col_widths = self.page_width / len(df.columns)
        for col in df.columns:
            self.cell(col_widths, 8, str(col), 1, 0, "C", True)
        self.ln()
        
        self.set_font("Arial", "", 8)
        self.set_text_color(*COLOR_TEXTO_CUERPO)
        for i, row in df.iterrows():
            fill = (i % 2 == 1)
            self.set_fill_color(*COLOR_GRIS_FONDO_FILA)
            for val in row:
                self.cell(col_widths, 7, str(val), 1, 0, "C", fill)
            self.ln()
        self.ln(5)

# --- FUNCIONES DE PROCESO ---

def procesar_flujo_rrhh(archivo, f_inicio, f_fin):
    # Cargar datos
    df_base = pd.read_excel(archivo, sheet_name='BaseQuery')
    df_activos_viejos = pd.read_excel(archivo, sheet_name='Activos')
    
    # Renombrar columnas de SAP
    mapping = {'Gr.prof.': 'Categoría', 'División de personal': 'Línea'}
    df_base.rename(columns=mapping, inplace=True)
    df_activos_viejos.rename(columns=mapping, inplace=True)

    try:
        df_co_manual = pd.read_excel(archivo, sheet_name='CO')
    except:
        df_co_manual = pd.DataFrame(columns=['Nº pers.', 'Apellido', 'Nombre de pila', 'Línea', 'Categoría', 'Desde', 'Motivo'])

    # Estandarizar
    for df in [df_base, df_activos_viejos, df_co_manual]:
        df['Nº pers.'] = df['Nº pers.'].astype(str).str.strip()

    # 1. Identificar Salidas por Comparación
    legajos_viejos = set(df_activos_viejos['Nº pers.'])
    legajos_activos_nuevos = set(df_base[df_base['Status ocupación'] == 'Activo']['Nº pers.'])
    ids_salidas = legajos_viejos - legajos_activos_nuevos

    # 2. Bajas Sistema (AJUSTE DE FECHA REAL)
    df_bajas = df_base[(df_base['Nº pers.'].isin(ids_salidas)) & (df_base['Status ocupación'] == 'Dado de baja')].copy()
    df_bajas['Desde'] = pd.to_datetime(df_bajas['Desde'])
    df_bajas['Fecha_Real'] = df_bajas['Desde'] - pd.Timedelta(days=1)
    df_bajas['Tipo'] = 'Baja'

    # 3. Cambios Organizativos (CO)
    ids_en_base = set(df_base['Nº pers.'])
    ids_co = ids_salidas - ids_en_base
    df_co_detectados = df_co_manual[df_co_manual['Nº pers.'].isin(ids_co)].copy()
    df_co_detectados['Fecha_Real'] = pd.to_datetime(df_co_detectados['Desde'])
    df_co_detectados['Tipo'] = 'Cambio Organizativo'
    if 'Motivo' in df_co_detectados.columns:
        df_co_detectados['Motivo de la medida'] = df_co_detectados['Motivo']
    else:
        df_co_detectados['Motivo de la medida'] = 'Reubicado'

    # 4. Unificar y Filtrar por el rango de reporte
    columnas = ['Nº pers.', 'Apellido', 'Nombre de pila', 'Línea', 'Categoría', 'Fecha_Real', 'Motivo de la medida', 'Tipo']
    df_all = pd.concat([df_bajas[columnas], df_co_detectados[columnas]], ignore_index=True)
    
    # Filtro de fecha para el reporte
    mask = (df_all['Fecha_Real'] >= pd.to_datetime(f_inicio)) & (df_all['Fecha_Real'] <= pd.to_datetime(f_fin))
    df_final = df_all[mask].copy()
    
    return df_final

# --- INTERFAZ STREAMLIT ---
st.set_page_config(page_title="Gestión de Bajas y CO", layout="wide")
st.title("📊 Análisis de Salidas: Bajas y Cambios Org.")

# Sidebar para filtros
st.sidebar.header("Configuración del Reporte")
f_inicio = st.sidebar.date_input("Fecha Inicio", datetime(2025, 6, 1))
f_fin = st.sidebar.date_input("Fecha Fin", datetime(2026, 1, 15))

archivo = st.file_uploader("Subir Excel (BaseQuery, Activos, CO)", type=['xlsx'])

if archivo:
    df_salidas = procesar_flujo_rrhh(archivo, f_inicio, f_fin)
    
    if not df_salidas.empty:
        st.subheader(f"Evolución del Período: {f_inicio.strftime('%d/%m/%Y')} al {f_fin.strftime('%d/%m/%Y')}")
        
        # Dashboard
        k1, k2, k3 = st.columns(3)
        k1.metric("Total Salidas", len(df_salidas))
        k2.metric("Bajas", len(df_salidas[df_salidas['Tipo'] == 'Baja']))
        k3.metric("C.O.", len(df_salidas[df_salidas['Tipo'] == 'Cambio Organizativo']))

        # Cuadro de Motivos Unificado
        st.write("### 📝 Motivos de Salida (Bajas + CO)")
        resumen_motivos = df_salidas.groupby(['Motivo de la medida', 'Tipo']).size().unstack(fill_value=0)
        resumen_motivos['Total'] = resumen_motivos.sum(axis=1)
        st.dataframe(resumen_motivos.sort_values('Total', ascending=False), use_container_width=True)

        # Gráfico Mensual
        df_salidas['Mes'] = df_salidas['Fecha_Real'].dt.strftime('%Y-%m')
        fig = px.bar(df_salidas.sort_values('Fecha_Real'), x='Mes', color='Tipo', barmode='group',
                     color_discrete_map={'Baja': '#ef553b', 'Cambio Organizativo': '#636efa'})
        st.plotly_chart(fig, use_container_width=True)

        # Detalle Nominal
        st.write("### 👥 Detalle de Personas")
        st.dataframe(df_salidas.sort_values('Fecha_Real'), hide_index=True)

        # Exportación PDF
        if st.button("📄 Exportar Reporte a PDF"):
            pdf = PDF(orientation='L', unit='mm', format='A4')
            pdf.report_title = f"Reporte de Bajas y C.O. ({f_inicio} a {f_fin})"
            pdf.add_page()
            
            # Tabla de Motivos
            pdf.draw_table("Resumen de Motivos", resumen_motivos.reset_index())
            
            # Tabla Nominal (solo columnas clave para que entre)
            pdf.draw_table("Detalle Nominal", df_salidas[['Nº pers.', 'Apellido', 'Línea', 'Fecha_Real', 'Motivo de la medida', 'Tipo']].astype(str))
            
            pdf_output = pdf.output(dest='S').encode('latin-1', 'replace')
            st.download_button("Descargar PDF", data=pdf_output, file_name="Reporte_Salidas.pdf", mime="application/pdf")
    else:
        st.warning("No se detectaron salidas para el rango seleccionado.")
