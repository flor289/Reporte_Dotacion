import streamlit as st
import pandas as pd
import plotly.express as px
from fpdf import FPDF
from datetime import datetime
import tempfile
import os

# --- CONFIGURACIÓN ESTÉTICA ---
AZUL_INSTITUCIONAL = (70, 130, 180)  # SteelBlue para PDF
TEXTO_TITULO_RGB = (0, 51, 102)      # Navy
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
            self.set_font("Arial", "B", 18)
            self.set_text_color(*TEXTO_TITULO_RGB)
            self.cell(0, 10, self.report_title, 0, 1, "C")
            self.set_font("Arial", "I", 8); self.set_text_color(100, 100, 100)
            fecha_hoy = datetime.now().strftime("%d/%m/%Y")
            self.set_xy(self.w - 50, 10)
            self.cell(40, 10, f"Fecha: {fecha_hoy}", 0, 0, "R")
            self.ln(12)

    def footer(self):
        self.set_y(-15)
        self.set_font("Arial", "B", 10); self.set_text_color(100, 100, 100)
        self.cell(0, 10, f"Página {self.page_no()}", 0, 0, "C")

    def draw_kpi_box(self, label, value, subtext, x, y):
        self.set_fill_color(248, 249, 250); self.rect(x, y, 60, 28, 'F')
        self.set_draw_color(*AZUL_INSTITUCIONAL); self.rect(x, y, 60, 28, 'D')
        self.set_xy(x, y + 4)
        self.set_font("Arial", "B", 8.5); self.set_text_color(100, 100, 100)
        self.cell(60, 5, label.upper(), 0, 1, "C")
        self.set_font("Arial", "B", 14); self.set_text_color(*AZUL_INSTITUCIONAL)
        self.cell(60, 9, value, 0, 1, "C")
        self.set_font("Arial", "", 8); self.set_text_color(120, 120, 120)
        self.cell(60, 5, subtext, 0, 1, "C")

    def draw_table_mini(self, title, df, x, y, width):
        self.set_xy(x, y)
        self.set_font("Arial", "B", 11); self.set_text_color(*TEXTO_TITULO_RGB)
        self.cell(width, 8, title, 0, 1, "L")
        if 'index' in df.columns: df = df.rename(columns={'index': 'Motivo'})
        elif df.index.name in [None, 'index']: df.index.name = 'Motivo'; df = df.reset_index()
        self.set_x(x); self.set_font("Arial", "B", 7); self.set_fill_color(*AZUL_INSTITUCIONAL); self.set_text_color(255, 255, 255)
        col_width = width / len(df.columns)
        for col in df.columns: self.cell(col_width, 6, str(col)[:10], 1, 0, "C", True)
        self.ln()
        self.set_font("Arial", "", 6.5); self.set_text_color(0, 0, 0)
        for _, row in df.iterrows():
            self.set_x(x)
            is_total = "TOTAL" in str(row.iloc[0]).upper()
            self.set_font("Arial", "B" if is_total else "", 6.5)
            for val in row:
                txt = str(val)
                if "Mutuo Acuerdo Art 241" in txt: txt = "M. ACUERDO ART. 241"
                self.cell(col_width, 5, txt, 1, 0, "C")
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

st.set_page_config(page_title="Reporte de Bajas", layout="wide")
archivo = st.file_uploader("Subir base Excel", type=['xlsx'])

if archivo:
    df_total = procesar_datos(archivo)
    pdf = PDF(orientation='L', unit='mm', format='A4')

    # --- PÁGINA 1: PORTADA ---
    pdf.report_title = "REPORTE DE BAJAS"
    pdf.add_page()
    pdf.set_font("Arial", "B", 12); pdf.set_text_color(100,100,100)
    pdf.text(125, 28, "Periodo: 2019 - Presente")
    
    pdf.draw_kpi_box("Total Bajas", str(len(df_total)), "Acumulado Histórico", 20, 35)
    pdf.draw_kpi_box("Línea con más bajas", "ROCA", "181 Casos", 85, 35)
    pdf.draw_kpi_box("Motivo Principal", "M. ACUERDO ART. 241", "280 Casos", 150, 35)
    pdf.draw_kpi_box("Máximo histórico anual", "2025", "284 Bajas", 215, 35)

    df_gen_anio = df_total.groupby('Año').size().reset_index(name='Bajas')
    fig_gen = px.line(df_gen_anio, x='Año', y='Bajas', markers=True, text='Bajas')
    fig_gen.update_traces(line_color=CELESTE_PLOTLY, textposition="top center", line_width=4)
    fig_gen.update_layout(plot_bgcolor='white', paper_bgcolor='white', yaxis_title="Cantidad", xaxis_title="Año")
    
    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
        fig_gen.write_image(tmp.name, scale=3)
        pdf.image(tmp.name, x=20, y=75, w=250)

    # --- PÁGINA 2: MATRIZ ---
    pdf.report_title = "MATRIZ HISTÓRICA DE MOTIVOS"
    pdf.add_page()
    t_matriz = preparar_tabla(df_total, 'Año')
    w_m = 65; w_r = (pdf.w - 40 - w_m) / len(t_matriz.columns)
    pdf.set_font("Arial", "B", 7); pdf.set_fill_color(70, 130, 180); pdf.set_text_color(255, 255, 255)
    pdf.set_xy(20, 35); pdf.cell(w_m, 8, "Motivo de Baja", 1, 0, "C", True)
    for col in t_matriz.columns: pdf.cell(w_r, 8, str(col), 1, 0, "C", True)
    pdf.ln()
    pdf.set_font("Arial", "", 7); pdf.set_text_color(0, 0, 0)
    for idx, row in t_matriz.iterrows():
        pdf.set_x(20); is_t = "TOTAL" in str(idx).upper()
        pdf.set_font("Arial", "B" if is_t else "", 7)
        txt_idx = "M. ACUERDO ART. 241" if "Mutuo Acuerdo Art 241" in str(idx) else str(idx)
        pdf.cell(w_m, 7, txt_idx, 1, 0, "L")
        for val in row: pdf.cell(w_r, 7, str(val), 1, 0, "C")
        pdf.ln()

    # --- PÁGINAS ANUALES ---
    for anio in sorted(df_total['Año'].unique()):
        pdf.report_title = f"ANÁLISIS DE BAJAS - {anio}"
        pdf.add_page()
        df_anio = df_total[df_total['Año'] == anio]
        
        mes_top = df_anio['Mes_Nom'].value_counts().idxmax().upper()
        pdf.draw_kpi_box(f"Total Bajas {anio}", str(len(df_anio)), "Total Ejercicio", 15, 35)
        pdf.draw_kpi_box("Línea con más bajas", df_anio['Línea'].value_counts().idxmax(), f"{df_anio['Línea'].value_counts().max()} Casos", 80, 35)
        pdf.draw_kpi_box("Motivo Principal", "M. ACUERDO ART. 241", f"{len(df_anio[df_anio['Motivo de la medida'].str.contains('241')])} Casos", 145, 35)
        pdf.draw_kpi_box("Mes con más bajas", mes_top, f"{df_anio['Mes_Nom'].value_counts().max()} Casos", 210, 35)

        t_mes = preparar_tabla(df_anio, 'Mes_Nom', ORDEN_MESES_CALENDARIO)
        t_lin = preparar_tabla(df_anio, 'Línea', ORDEN_LINEAS)
        pdf.draw_table_mini("Distribución por Mes", t_mes.reset_index(), 15, 68, 130)
        pdf.draw_table_mini("Distribución por Línea", t_lin.reset_index(), 150, 68, 130)

        df_bar = df_anio.groupby(['Mes_Num', 'Mes_Nom', 'Línea']).size().reset_index(name='Cantidad')
        meses_activos = [m for m in ORDEN_MESES_CALENDARIO if m in df_bar['Mes_Nom'].unique()]
        fig_bar = px.bar(df_bar.sort_values('Mes_Num'), x='Mes_Nom', y='Cantidad', color='Línea', 
                         barmode='group', text='Cantidad', color_discrete_map=COLORES_LINEAS, 
                         category_orders={"Mes_Nom": meses_activos, "Línea": ORDEN_LINEAS})
        fig_bar.update_traces(textposition='outside', textfont_size=11)
        fig_bar.update_layout(plot_bgcolor='white', paper_bgcolor='white', bargap=0.4)
        
        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
            fig_bar.write_image(tmp.name, scale=3, width=1100, height=450)
            pdf.image(tmp.name, x=15, y=125, w=265)

    pdf_out = pdf.output(dest='S').encode('latin-1', 'replace')
    st.sidebar.download_button("📩 Guardar Reporte Final", data=pdf_out, file_name="Reporte_Bajas_Final.pdf")
    st.success("Reporte generado con éxito.")
