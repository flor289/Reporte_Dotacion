import streamlit as st
import pandas as pd
import plotly.express as px
from fpdf import FPDF
from datetime import datetime
import tempfile
import os

# --- CONFIGURACIÓN DE COLORES ---
AZUL_INSTITUCIONAL = "#4682B4"  # Azul acero para encabezados
TEXTO_TITULO_RGB = (0, 51, 102)
COLORES_LINEAS = {
    "ROCA": "#3A70A9", "SARMIENTO": "#8AA0B9", "BELGRANO SUR": "#FDC84A",
    "SAN MARTIN": "#CD5055", "MITRE": "#5F8751", "REGIONALES": "#7B6482", "CENTRAL": "#808080"
}
MESES_ES = {1: 'Ene', 2: 'Feb', 3: 'Mar', 4: 'Abr', 5: 'May', 6: 'Jun', 
            7: 'Jul', 8: 'Ago', 9: 'Sep', 10: 'Oct', 11: 'Nov', 12: 'Dic'}

class PDF(FPDF):
    def header(self):
        if hasattr(self, 'report_title') and self.report_title:
            self.set_font("Arial", "B", 16) # Tamaño grande para títulos de gráfico/reporte
            self.set_text_color(*TEXTO_TITULO_RGB)
            self.cell(0, 12, self.report_title, 0, 1, "C")
            self.ln(5)

    def footer(self):
        self.set_y(-15)
        self.set_font("Arial", "", 10)
        self.set_text_color(100, 100, 100)
        self.cell(0, 10, f"{self.page_no()}", 0, 0, "C")

    def draw_table(self, title, df):
        self.set_font("Arial", "B", 16) # Título de tabla igual al de gráfico
        self.set_text_color(*TEXTO_TITULO_RGB)
        self.cell(0, 10, title, ln=True)
        
        self.set_font("Arial", "B", 8)
        self.set_fill_color(70, 130, 180) # Azul institucional encabezado
        self.set_text_color(255, 255, 255) # Letra blanca
        
        col_widths = (self.w - 20) / len(df.columns)
        for col in df.columns:
            self.cell(col_widths, 8, str(col), 1, 0, "C", True)
        self.ln()
        
        self.set_font("Arial", "", 8)
        self.set_text_color(0, 0, 0)
        for i, row in df.iterrows():
            # Si es la fila de TOTAL, poner en negrita
            if "TOTAL" in str(row.iloc[0]).upper(): self.set_font("Arial", "B", 8)
            else: self.set_font("Arial", "", 8)
            for val in row:
                self.cell(col_widths, 7, str(val), 1, 0, "C") # Datos centrados
            self.ln()
        self.ln(5)

def procesar_datos(archivo):
    df_base = pd.read_excel(archivo, sheet_name='BaseQuery')
    mapping = {'Gr.prof.': 'Categoría', 'División de personal': 'Línea', 'Division de personal': 'Línea'}
    df_base.rename(columns=mapping, inplace=True)
    df_base['Línea'] = df_base['Línea'].astype(str).str.upper().str.strip()
    df_bajas = df_base[df_base['Status ocupación'] == 'Dado de baja'].copy()
    df_bajas['Desde'] = pd.to_datetime(df_bajas['Desde'])
    df_bajas['Fecha_Real'] = df_bajas['Desde'] - pd.Timedelta(days=1)
    df_bajas = df_bajas[df_bajas['Fecha_Real'].dt.year >= 2019]
    df_bajas['Año'] = df_bajas['Fecha_Real'].dt.year
    df_bajas['Mes_Num'] = df_bajas['Fecha_Real'].dt.month
    df_bajas['Mes_Nom'] = df_bajas['Mes_Num'].map(MESES_ES)
    return df_bajas

st.set_page_config(page_title="Reporte Bajas", layout="wide")
archivo = st.file_uploader("Subir Archivo Excel", type=['xlsx'])

if archivo:
    df_total = procesar_datos(archivo)
    pdf = PDF(orientation='L', unit='mm', format='A4')

    # --- 1. RESUMEN GENERAL ---
    st.markdown(f"<h1 style='text-align: center; color: {AZUL_INSTITUCIONAL};'>Resumen General de Bajas</h1>", unsafe_allow_html=True)
    
    df_gen_anio = df_total.groupby('Año').size().reset_index(name='Bajas')
    fig_gen = px.line(df_gen_anio, x='Año', y='Bajas', markers=True, text='Bajas', title="Evolución Anual de Bajas")
    fig_gen.update_traces(line_color=AZUL_INSTITUCIONAL, textposition="top center", line_width=4)
    fig_gen.update_layout(title_font_size=24, xaxis_title="Año", yaxis_title="Cantidad")
    st.plotly_chart(fig_gen, use_container_width=True)

    t_gen = df_total.pivot_table(index='Motivo de la medida', columns='Año', values='Nº pers.', aggfunc='count', fill_value=0)
    t_gen.loc['TOTAL GENERAL'] = t_gen.sum()
    t_gen['Total'] = t_gen.sum(axis=1)
    st.subheader("Motivos de Baja por Año")
    st.dataframe(t_gen.style.set_properties(**{'text-align': 'center'}), use_container_width=True)

    # --- 2. DESGLOSE POR AÑO (ORDEN 2019 -> ACTUAL) ---
    años = sorted(df_total['Año'].unique())
    for anio in años:
        st.markdown("---")
        titulo_anio = f"REPORTE ANUAL DE BAJAS - {anio}"
        st.markdown(f"<h1 style='text-align: center; color: {AZUL_INSTITUCIONAL};'>{titulo_anio}</h1>", unsafe_allow_html=True)
        df_anio = df_total[df_total['Año'] == anio]

        # Tablas una debajo de otra
        st.markdown(f"### Motivos de Baja por Mes - {anio}")
        t_mes = df_anio.pivot_table(index='Motivo de la medida', columns='Mes_Nom', values='Nº pers.', aggfunc='count', fill_value=0)
        cols_m = [m for m in MESES_ES.values() if m in t_mes.columns]
        t_mes = t_mes[cols_m]
        t_mes.loc['TOTAL'] = t_mes.sum()
        t_mes['Total Anual'] = t_mes.sum(axis=1)
        st.dataframe(t_mes.style.set_properties(**{'text-align': 'center'}), use_container_width=True)

        st.markdown(f"### Motivos de Baja por Línea - {anio}")
        t_lin = df_anio.pivot_table(index='Motivo de la medida', columns='Línea', values='Nº pers.', aggfunc='count', fill_value=0)
        t_lin.loc['TOTAL'] = t_lin.sum()
        t_lin['Total Anual'] = t_lin.sum(axis=1)
        st.dataframe(t_lin.style.set_properties(**{'text-align': 'center'}), use_container_width=True)

        # Gráfico Mensual (Mejorado para que no sea feo)
        df_evol = df_anio.groupby(['Mes_Num', 'Mes_Nom', 'Línea']).size().reset_index(name='Cantidad')
        fig_m = px.line(df_evol.sort_values('Mes_Num'), x='Mes_Nom', y='Cantidad', color='Línea', 
                        markers=True, text='Cantidad', title="Evolución Mensual de Bajas por Línea",
                        color_discrete_map=COLORES_LINEAS)
        fig_m.update_traces(textposition="top center", line_width=3)
        fig_m.update_layout(title_font_size=24, xaxis_title="Mes", yaxis_title="Cantidad", hovermode="x unified")
        st.plotly_chart(fig_m, use_container_width=True)

        # PDF LOGIC
        pdf.report_title = titulo_anio
        pdf.add_page()
        pdf.draw_table("Motivos de Baja por Mes", t_mes.reset_index())
        pdf.draw_table("Motivos de Baja por Línea", t_lin.reset_index())
        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
            fig_m.write_image(tmp.name)
            pdf.image(tmp.name, x=10, y=pdf.get_y() + 10, w=270)

    pdf_data = pdf.output(dest='S').encode('latin-1', 'replace')
    st.sidebar.download_button("Descargar PDF Completo", data=pdf_data, file_name="Reporte_Bajas_Final.pdf")
