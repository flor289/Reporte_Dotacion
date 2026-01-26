import streamlit as st
import pandas as pd
import plotly.express as px
from fpdf import FPDF
import tempfile
import os

# --- CONFIGURACIÓN ESTÉTICA ---
AZUL_INSTITUCIONAL = "#4682B4" 
CELESTE_INSTITUCIONAL = "#7dbad2"
TEXTO_TITULO_RGB = (0, 51, 102)

# Orden y colores institucionales
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
        if hasattr(self, 'report_title') and self.report_title:
            self.set_font("Arial", "B", 16)
            self.set_text_color(*TEXTO_TITULO_RGB)
            self.cell(0, 12, self.report_title, 0, 1, "C")
            self.ln(5)

    def footer(self):
        self.set_y(-15)
        self.set_font("Arial", "", 10)
        self.set_text_color(100, 100, 100)
        self.cell(0, 10, f"{self.page_no()}", 0, 0, "C")

    def check_page_break(self, height_needed):
        if self.get_y() + height_needed > 180:
            self.add_page()
            return True
        return False

    def draw_table(self, title, df):
        # Renombrar index a Motivo de Baja
        if 'index' in df.columns:
            df = df.rename(columns={'index': 'Motivo de Baja'})
        elif df.index.name is None or df.index.name == 'index':
            df = df.reset_index().rename(columns={'index': 'Motivo de Baja'})

        self.set_font("Arial", "B", 16)
        self.set_text_color(*TEXTO_TITULO_RGB)
        self.cell(0, 10, title, ln=True)
        
        w_motivo = 65
        w_resto = (self.w - 20 - w_motivo) / (len(df.columns) - 1)
        
        self.set_font("Arial", "B", 7)
        self.set_fill_color(70, 130, 180) 
        self.set_text_color(255, 255, 255) 
        
        for i, col in enumerate(df.columns):
            w = w_motivo if i == 0 else w_resto
            self.cell(w, 8, str(col), 1, 0, "C", True)
        self.ln()
        
        self.set_font("Arial", "", 7); self.set_text_color(0, 0, 0)
        for _, row in df.iterrows():
            is_total = "TOTAL" in str(row.iloc[0]).upper()
            self.set_font("Arial", "B" if is_total else "", 7)
            for i, val in enumerate(row):
                w = w_motivo if i == 0 else w_resto
                texto = str(val)
                if i == 0 and len(texto) > 45: texto = texto[:42] + "..."
                self.cell(w, 7, texto, 1, 0, "C") 
            self.ln()
        self.ln(5)

def preparar_tabla_final(df, index_c, order_c=None):
    t = df.pivot_table(index='Motivo de la medida', columns=index_c, values='Nº pers.', aggfunc='count', fill_value=0)
    if order_c: t = t[[c for c in order_c if c in t.columns]]
    t['Total Anual'] = t.sum(axis=1)
    t = t.sort_values('Total Anual', ascending=False)
    # Total al final
    f_t = t.sum().to_frame().T
    f_t.index = ['TOTAL']
    return pd.concat([t, f_t]).replace(0, '-')

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

# --- APP ---
st.set_page_config(page_title="Reporte Bajas", layout="wide")
archivo = st.file_uploader("Subir Excel", type=['xlsx'])

if archivo:
    df_total = procesar_datos(archivo)
    pdf = PDF(orientation='L', unit='mm', format='A4')

    # --- 1. RESUMEN GENERAL ---
    st.markdown(f"<h1 style='text-align: center; color: {AZUL_INSTITUCIONAL};'>Resumen General de Bajas</h1>", unsafe_allow_html=True)
    df_gen_anio = df_total.groupby('Año').size().reset_index(name='Bajas')
    fig_gen = px.line(df_gen_anio, x='Año', y='Bajas', markers=True, text='Bajas', title="Evolución Anual de Bajas")
    fig_gen.update_traces(line_color=CELESTE_INSTITUCIONAL, textposition="top center", line_width=4, marker=dict(size=12))
    fig_gen.update_layout(
        title_font_size=24, plot_bgcolor='white', paper_bgcolor='white',
        yaxis=dict(tickformat='d', nticks=10, showgrid=True, gridcolor='#F0F0F0', title="Cantidad"),
        xaxis=dict(showgrid=True, gridcolor='#F0F0F0', dtick=1)
    )
    st.plotly_chart(fig_gen, use_container_width=True)

    t_gen = preparar_tabla_final(df_total, 'Año')
    st.markdown("### Motivos de Baja por Año")
    st.dataframe(t_gen.style.set_properties(**{'text-align': 'center'}), use_container_width=True)

    pdf.report_title = "RESUMEN GENERAL DE BAJAS (2019 - Presente)"
    pdf.add_page()
    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_gen:
        fig_gen.write_image(tmp_gen.name, scale=3, width=1200, height=600)
        pdf.image(tmp_gen.name, x=15, y=35, w=260)
    pdf.add_page(); pdf.draw_table("Motivos de Baja por Año", t_gen)

    # --- 2. DESGLOSE POR AÑO ---
    años = sorted(df_total['Año'].unique())
    for anio in años:
        st.markdown("---")
        tit_label = f"REPORTE ANUAL DE BAJAS - {anio}"
        st.markdown(f"<h1 style='text-align: center; color: {AZUL_INSTITUCIONAL};'>{tit_label}</h1>", unsafe_allow_html=True)
        df_anio = df_total[df_total['Año'] == anio]

        t_mes = preparar_tabla_final(df_anio, 'Mes_Nom', ORDEN_MESES_CALENDARIO)
        t_lin = preparar_tabla_final(df_anio, 'Línea', ORDEN_LINEAS)

        st.markdown(f"### Motivos de Baja por Mes")
        st.dataframe(t_mes.style.set_properties(**{'text-align': 'center'}), use_container_width=True)
        st.markdown(f"### Motivos de Baja por Línea")
        st.dataframe(t_lin.style.set_properties(**{'text-align': 'center'}), use_container_width=True)

        # Lógica para no mostrar meses sin bajas
        df_bar = df_anio.groupby(['Mes_Num', 'Mes_Nom', 'Línea']).size().reset_index(name='Cantidad')
        meses_con_datos = [m for m in ORDEN_MESES_CALENDARIO if m in df_bar['Mes_Nom'].unique()]
        
        fig_bar = px.bar(df_bar.sort_values('Mes_Num'), x='Mes_Nom', y='Cantidad', color='Línea', 
                         barmode='group', text='Cantidad', title="Evolución Mensual de Bajas por Línea",
                         color_discrete_map=COLORES_LINEAS, 
                         category_orders={"Mes_Nom": meses_con_datos, "Línea": ORDEN_LINEAS})
        
        max_v = df_bar['Cantidad'].max()
        fig_bar.update_layout(
            title_font_size=24, plot_bgcolor='white', paper_bgcolor='white',
            yaxis=dict(tickformat='d', nticks=10, range=[0, max_v + 1] if max_v < 4 else None, title="Cantidad"),
            bargap=0.8
        )
        st.plotly_chart(fig_bar, use_container_width=True)

        pdf.report_title = tit_label; pdf.add_page()
        pdf.draw_table("Motivos de Baja por Mes", t_mes)
        pdf.check_page_break(50)
        pdf.draw_table("Motivos de Baja por Línea", t_lin)
        pdf.check_page_break(110)
        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_anio:
            fig_bar.write_image(tmp_anio.name, scale=3, width=1200, height=600)
            pdf.image(tmp_anio.name, x=15, y=pdf.get_y() + 5, w=260)

    pdf_out = pdf.output(dest='S').encode('latin-1', 'replace')
    st.sidebar.download_button("📩 Descargar PDF Final", data=pdf_out, file_name="Reporte_Bajas_Final.pdf")
