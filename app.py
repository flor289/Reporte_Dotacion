import streamlit as st
import pandas as pd
import plotly.express as px
from fpdf import FPDF
from datetime import datetime
import tempfile
import os

# --- MAPEO DE COLORES INSTITUCIONALES ---
COLORES_LINEAS = {
    "Línea Roca": "#3A70A9",
    "Línea Sarmiento": "#8AA0B9",
    "Línea Belgrano Sur": "#FDC84A",
    "Línea San Martín": "#CD5055",
    "Línea Mitre y Tren de la Costa": "#5F8751",
    "Línea Mitre": "#5F8751", # Por si varía el nombre
    "Regionales": "#7B6482",
    "Central": "#808080"
}

COLOR_TEXTO_TITULO = (0, 51, 102)
MESES_ES = {1: 'Ene', 2: 'Feb', 3: 'Mar', 4: 'Abr', 5: 'May', 6: 'Jun', 
            7: 'Jul', 8: 'Ago', 9: 'Sep', 10: 'Oct', 11: 'Nov', 12: 'Dic'}

class PDF(FPDF):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.report_title = ""

    def header(self):
        if self.report_title:
            self.set_font("Arial", "B", 14)
            self.set_text_color(*COLOR_TEXTO_TITULO)
            self.cell(0, 10, self.report_title, 0, 1, "C")
            self.ln(5)

    def footer(self):
        self.set_y(-15)
        self.set_font("Arial", "", 9)
        self.cell(0, 10, f"{self.page_no()}", 0, 0, "C")

    def draw_table(self, title, df):
        if df.empty: return
        self.set_font("Arial", "B", 10)
        self.set_text_color(*COLOR_TEXTO_TITULO)
        self.cell(0, 8, title, ln=True)
        self.set_font("Arial", "B", 8)
        self.set_fill_color(240, 242, 246)
        self.set_text_color(0, 0, 0)
        
        col_widths = (self.w - 20) / len(df.columns)
        for col in df.columns:
            self.cell(col_widths, 7, str(col), 1, 0, "C", True)
        self.ln()
        
        self.set_font("Arial", "", 8)
        for i, row in df.iterrows():
            for val in row:
                self.cell(col_widths, 6, str(val), 1, 0, "C")
            self.ln()
        self.ln(5)

def procesar_datos(archivo):
    df_base = pd.read_excel(archivo, sheet_name='BaseQuery')
    mapping = {'Gr.prof.': 'Categoría', 'División de personal': 'Línea', 'Division de personal': 'Línea'}
    df_base.rename(columns=mapping, inplace=True)
    
    df_bajas = df_base[df_base['Status ocupación'] == 'Dado de baja'].copy()
    df_bajas['Desde'] = pd.to_datetime(df_bajas['Desde'])
    # REGLA: Fecha Real = Desde - 1 día
    df_bajas['Fecha_Real'] = df_bajas['Desde'] - pd.Timedelta(days=1)
    
    df_bajas = df_bajas[df_bajas['Fecha_Real'].dt.year >= 2019]
    df_bajas['Año'] = df_bajas['Fecha_Real'].dt.year
    df_bajas['Mes_Num'] = df_bajas['Fecha_Real'].dt.month
    df_bajas['Mes_Nom'] = df_bajas['Mes_Num'].map(MESES_ES)
    df_bajas['Mes_Anio'] = df_bajas['Mes_Nom'] + "-" + df_bajas['Año'].astype(str).str[-2:]
    return df_bajas

st.set_page_config(page_title="Reporte de Bajas RRHH", layout="wide")
archivo = st.file_uploader("Subir base de datos (Excel)", type=['xlsx'])

if archivo:
    df_total = procesar_datos(archivo)
    pdf = PDF(orientation='L', unit='mm', format='A4')
    
    # --- RESUMEN GENERAL (Sin "Vista Histórica") ---
    st.title("Resumen General de Bajas")
    
    # 1. Gráfico Evolutivo Anual
    df_gen_anio = df_total.groupby('Año').size().reset_index(name='Bajas')
    fig_gen = px.line(df_gen_anio, x='Año', y='Bajas', markers=True, text='Bajas', title="Evolución Anual de Bajas")
    fig_gen.update_traces(textposition="top center", line_color="#003366")
    st.plotly_chart(fig_gen, use_container_width=True)
    
    # 2. Tabla Motivos de Baja por Año
    st.subheader("Motivos de Baja por Año")
    t_motivos_anio = df_total.pivot_table(index='Motivo de la medida', columns='Año', values='Nº pers.', aggfunc='count', fill_value=0)
    t_motivos_anio['Total'] = t_motivos_anio.sum(axis=1)
    t_motivos_anio = t_motivos_anio.sort_values('Total', ascending=False).replace(0, '-')
    st.dataframe(t_motivos_anio, use_container_width=True)

    # 3. Evolutivo de Líneas Anual (General)
    df_evol_l_gen = df_total.groupby(['Año', 'Línea']).size().reset_index(name='Cant')
    fig_l_gen = px.line(df_evol_l_gen, x='Año', y='Cant', color='Línea', markers=True, text='Cant',
                        title="Evolución Anual de Bajas por Línea", color_discrete_map=COLORES_LINEAS)
    fig_l_gen.update_traces(textposition="top center")
    st.plotly_chart(fig_l_gen, use_container_width=True)

    # PDF: Primera Página
    pdf.report_title = "RESUMEN GENERAL DE BAJAS (2019 - Presente)"
    pdf.add_page()
    pdf.draw_table("Motivos de Baja por Año", t_motivos_anio.reset_index())
    
    # --- SECCIÓN POR AÑO ---
    años = sorted(df_total['Año'].unique(), reverse=True)
    for anio in años:
        st.markdown("---")
        st.header(f"REPORTE ANUAL DE BAJAS - {anio}")
        df_anio = df_total[df_total['Año'] == anio]
        
        # 1. Motivos de Baja por Mes
        t_mes = df_anio.pivot_table(index='Motivo de la medida', columns='Mes_Anio', values='Nº pers.', aggfunc='count', fill_value=0)
        # Orden cronológico de meses
        cols_m = sorted(t_mes.columns, key=lambda x: list(MESES_ES.values()).index(x.split('-')[0]))
        t_mes = t_mes[cols_m]
        t_mes['Total'] = t_mes.sum(axis=1)
        t_mes = t_mes.sort_values('Total', ascending=False).replace(0, '-')
        
        # 2. Motivos de Baja por Línea
        t_linea = df_anio.pivot_table(index='Motivo de la medida', columns='Línea', values='Nº pers.', aggfunc='count', fill_value=0)
        t_linea['Total'] = t_linea.sum(axis=1)
        t_linea = t_linea.sort_values('Total', ascending=False).replace(0, '-')
        
        st.subheader("Tablas de Análisis")
        c1, c2 = st.columns(2)
        c1.write("**Motivos por Mes**")
        c1.dataframe(t_mes)
        c2.write("**Motivos por Línea**")
        c2.dataframe(t_linea)

        # 3. Evolución Mensual por Línea
        df_evol_m = df_anio.groupby(['Mes_Num', 'Mes_Nom', 'Línea']).size().reset_index(name='Cant')
        fig_m = px.line(df_evol_m.sort_values('Mes_Num'), x='Mes_Nom', y='Cant', color='Línea', 
                        title="Evolución Mensual de Bajas por Línea", markers=True, text='Cant',
                        color_discrete_map=COLORES_LINEAS)
        fig_m.update_traces(textposition="top center")
        st.plotly_chart(fig_m, use_container_width=True)

        # PDF: Hoja Anual
        pdf.report_title = f"REPORTE ANUAL DE BAJAS - {anio}"
        pdf.add_page()
        pdf.draw_table("Motivos de Baja por Mes", t_mes.reset_index())
        pdf.draw_table("Motivos de Baja por Línea", t_linea.reset_index())
        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
            fig_m.write_image(tmp.name)
            pdf.image(tmp.name, x=10, y=pdf.get_y() + 5, w=270)

    # Descarga
    pdf_out = pdf.output(dest='S').encode('latin-1', 'replace')
    st.sidebar.download_button("📩 Descargar Reporte Completo PDF", data=pdf_out, file_name="Reporte_Bajas_Trenes.pdf")
