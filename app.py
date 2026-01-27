import streamlit as st
import pandas as pd
import plotly.express as px
from fpdf import FPDF
from datetime import datetime
import tempfile
import os

# --- CONFIGURACIÓN ESTÉTICA ---
AZUL_INSTITUCIONAL = (70, 130, 180) 
TEXTO_TITULO_RGB = (0, 51, 102)
CELESTE_PLOTLY = "#7dbad2"

ORDEN_LINEAS = ["ROCA", "MITRE", "SARMIENTO", "REGIONALES", "SAN MARTIN", "CENTRAL", "BELGRANO SUR"]
COLORES_LINEAS = {
    "ROCA": "#3A70A9", "SARMIENTO": "#8AA0B9", "BELGRANO SUR": "#FDC84A",
    "SAN MARTIN": "#CD5055", "MITRE": "#5F8751", "REGIONALES": "#7B6482", "CENTRAL": "#808080"
}
ORDEN_MESES_CALENDARIO = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic']
MESES_ES = {1: 'Ene', 2: 'Feb', 3: 'Mar', 4: 'Abr', 5: 'May', 6: 'Jun', 
            7: 'Jul', 8: 'Ago', 9: 'Sep', 10: 'Oct', 11: 'Nov', 12: 'Dic'}

class PDF(FPDF):
    def header(self):
        if hasattr(self, 'report_title'):
            # Título más pequeño y compacto
            self.set_font("Arial", "B", 14) 
            self.set_text_color(*TEXTO_TITULO_RGB)
            self.set_y(8)
            self.cell(0, 8, self.report_title, 0, 1, "C")
            
            self.set_font("Arial", "I", 8); self.set_text_color(150, 150, 150)
            self.set_xy(self.w - 45, 8)
            self.cell(35, 8, f"Fecha: {datetime.now().strftime('%d/%m/%Y')}", 0, 0, "R")

    def footer(self):
        self.set_y(-12)
        self.set_font("Arial", "B", 9); self.set_text_color(180, 180, 180)
        self.cell(0, 8, f"Página {self.page_no()}", 0, 0, "C")

    def draw_kpi_box(self, label, value, subtext, x, y):
        # Caja más pequeña (50x20)
        self.set_fill_color(248, 249, 250); self.rect(x, y, 50, 20, 'F')
        self.set_draw_color(*AZUL_INSTITUCIONAL); self.rect(x, y, 50, 20, 'D')
        
        self.set_xy(x, y + 2)
        self.set_font("Arial", "B", 7); self.set_text_color(100, 100, 100)
        self.cell(50, 4, label.upper(), 0, 1, "C")
        
        self.set_font("Arial", "B", 11); self.set_text_color(*AZUL_INSTITUCIONAL)
        self.set_x(x); self.cell(50, 7, value, 0, 1, "C")
        
        self.set_font("Arial", "", 6); self.set_text_color(120, 120, 120)
        self.set_x(x); self.cell(50, 4, subtext, 0, 1, "C")

    def draw_table_mini(self, title, df, x, y, total_w):
        self.set_xy(x, y)
        self.set_font("Arial", "B", 10); self.set_text_color(*TEXTO_TITULO_RGB)
        self.cell(total_w, 6, title, 0, 1, "L")
        
        if 'index' in df.columns: df = df.rename(columns={'index': 'Motivo'})
        elif df.index.name in [None, 'index']: df.index.name = 'Motivo'; df = df.reset_index()

        self.set_x(x); self.set_font("Arial", "B", 7); self.set_fill_color(*AZUL_INSTITUCIONAL); self.set_text_color(255, 255, 255)
        
        w_motivo = 40 if total_w > 120 else 30
        w_resto = (total_w - w_motivo) / (len(df.columns) - 1)
        
        for i, col in enumerate(df.columns):
            self.cell(w_motivo if i==0 else w_resto, 5, str(col), 1, 0, "C", True)
        self.ln()
        
        self.set_font("Arial", "", 6.5); self.set_text_color(0, 0, 0)
        for _, row in df.iterrows():
            self.set_x(x)
            is_total = "TOTAL" in str(row.iloc[0]).upper()
            self.set_font("Arial", "B" if is_total else "", 6.5)
            for i, val in enumerate(row):
                txt = str(val)
                if "Mutuo Acuerdo Art 241" in txt: txt = "M. ACUERDO ART. 241"
                self.cell(w_motivo if i==0 else w_resto, 4.5, txt, 1, 0, "C")
            self.ln()

def preparar_tabla(df, index_col, order=None):
    t = df.pivot_table(index='Motivo de la medida', columns=index_col, values='Nº pers.', aggfunc='count', fill_value=0)
    if order: t = t[[c for c in order if c in t.columns]]
    t['Total'] = t.sum(axis=1)
    t = t.sort_values('Total', ascending=False)
    f_total = t.sum().to_frame().T
    f_total.index = ['TOTAL']
    return pd.concat([t, f_total]).replace(0, '-')

def procesar_datos(archivo):
    df_base = pd.read_excel(archivo, sheet_name='BaseQuery')
    df_base.rename(columns={'Gr.prof.': 'Categoría', 'División de personal': 'Línea', 'Division de personal': 'Línea'}, inplace=True)
    df_base['Línea'] = df_base['Línea'].astype(str).str.upper().str.strip()
    df_bajas = df_base[df_base['Status ocupación'] == 'Dado de baja'].copy()
    df_bajas['Desde'] = pd.to_datetime(df_bajas['Desde'])
    df_bajas['Fecha_Real'] = df_bajas['Desde'] - pd.Timedelta(days=1)
    df_bajas = df_bajas[df_bajas['Fecha_Real'].dt.year >= 2019]
    df_bajas['Año'] = df_bajas['Fecha_Real'].dt.year
    df_bajas['Mes_Num'] = df_bajas['Fecha_Real'].dt.month
    df_bajas['Mes_Nom'] = df_bajas['Mes_Num'].map(MESES_ES)
    return df_bajas

# --- APP ---
st.set_page_config(page_title="Gestión de Bajas", layout="wide")
archivo = st.file_uploader("Subir Archivo Excel", type=['xlsx'])

if archivo:
    df_total = procesar_datos(archivo)
    pdf = PDF(orientation='L', unit='mm', format='A4')

    # --- PORTADA ---
    pdf.report_title = "REPORTE DE BAJAS"
    pdf.add_page()
    
    # KPIs Compactos y Centrados
    x_start = (297 - (4 * 50 + 3 * 10)) / 2 # Centrado dinámico
    pdf.draw_kpi_box("Total Bajas", str(len(df_total)), "2019 - Presente", x_start, 22)
    pdf.draw_kpi_box("Línea c/ más bajas", "ROCA", "181 Casos", x_start + 60, 22)
    pdf.draw_kpi_box("Motivo Principal", "M. ACUERDO ART. 241", "280 Casos", x_start + 120, 22)
    pdf.draw_kpi_box("Máximo Anual", "2025", "284 Bajas", x_start + 180, 22)

    df_gen_anio = df_total.groupby('Año').size().reset_index(name='Bajas')
    fig_gen = px.line(df_gen_anio, x='Año', y='Bajas', markers=True, text='Bajas')
    fig_gen.update_traces(line_color=CELESTE_PLOTLY, textposition="top center", line_width=3)
    # Forzar todos los años en eje X
    fig_gen.update_layout(plot_bgcolor='white', paper_bgcolor='white', 
                          yaxis=dict(dtick=50, tickformat='d', title="Cant."),
                          xaxis=dict(dtick=1, title="Año"), height=400)
    
    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
        fig_gen.write_image(tmp.name, scale=3)
        pdf.image(tmp.name, x=20, y=50, w=257, h=130)

    # --- MATRIZ ---
    pdf.report_title = "MATRIZ HISTÓRICA DE MOTIVOS"
    pdf.add_page()
    t_matriz = preparar_tabla(df_total, 'Año')
    pdf.draw_table_mini("Consolidado 2019-2026", t_matriz.reset_index(), 15, 22, 267)

    # --- ANÁLISIS POR AÑO ---
    for anio in sorted(df_total['Año'].unique()):
        pdf.report_title = f"ANÁLISIS DE BAJAS - {anio}"
        pdf.add_page()
        df_anio = df_total[df_total['Año'] == anio]
        
        # KPIs Anuales (y=22)
        m_top = df_anio['Mes_Nom'].value_counts().idxmax().upper()
        pdf.draw_kpi_box(f"Total Bajas {anio}", str(len(df_anio)), "Ejercicio Anual", x_start, 22)
        pdf.draw_kpi_box("Línea Crítica", df_anio['Línea'].value_counts().idxmax(), "Máxima Rotación", x_start + 60, 22)
        pdf.draw_kpi_box("Motivo Principal", "M. ACUERDO ART. 241", f"{len(df_anio[df_anio['Motivo de la medida'].str.contains('241')])} Casos", x_start + 120, 22)
        pdf.draw_kpi_box("Mes Crítico", m_top, "Pico de Bajas", x_start + 180, 22)

        t_mes = preparar_tabla(df_anio, 'Mes_Nom', ORDEN_MESES_CALENDARIO)
        t_lin = preparar_tabla(df_anio, 'Línea', ORDEN_LINEAS)
        
        # Tablas (y=45) con prioridad
        pdf.draw_table_mini("Distribución Mensual", t_mes.reset_index(), 15, 45, 130)
        pdf.draw_table_mini("Distribución por Línea", t_lin.reset_index(), 150, 45, 130)

        # Gráfico Normal (y=115) para no pisar el pie
        df_bar = df_anio.groupby(['Mes_Num', 'Mes_Nom', 'Línea']).size().reset_index(name='Cantidad')
        meses_activos = [m for m in ORDEN_MESES_CALENDARIO if m in df_bar['Mes_Nom'].unique()]
        fig_bar = px.bar(df_bar.sort_values('Mes_Num'), x='Mes_Nom', y='Cantidad', color='Línea', 
                         barmode='group', text='Cantidad', color_discrete_map=COLORES_LINEAS, 
                         category_orders={"Mes_Nom": meses_activos, "Línea": ORDEN_LINEAS})
        fig_bar.update_traces(textposition='outside', textfont_size=9)
        fig_bar.update_layout(plot_bgcolor='white', paper_bgcolor='white', bargap=0.3,
                              yaxis=dict(dtick=1, tickformat='d', title="Cant."), margin=dict(t=20, b=20))
        
        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
            fig_bar.write_image(tmp.name, scale=3, width=1100, height=300) # Más bajo
            pdf.image(tmp.name, x=20, y=115, w=257, h=75)

    pdf_out = pdf.output(dest='S').encode('latin-1', 'replace')
    st.sidebar.download_button("📩 Guardar Reporte Final", data=pdf_out, file_name="Reporte_Bajas_Final.pdf")


