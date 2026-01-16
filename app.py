import streamlit as st
import pandas as pd
import plotly.express as px
from fpdf import FPDF
from datetime import datetime, timedelta
import io
import tempfile
import os

# --- CONFIGURACIÓN Y ESTILOS ---
COLOR_AZUL_INSTITUCIONAL = (4, 118, 208)
COLOR_TEXTO_TITULO = (0, 51, 102)

# Mapeo de meses a español
MESES_ES = {
    1: 'Ene', 2: 'Feb', 3: 'Mar', 4: 'Abr', 5: 'May', 6: 'Jun',
    7: 'Jul', 8: 'Ago', 9: 'Sep', 10: 'Oct', 11: 'Nov', 12: 'Dic'
}

def estilo_kpi_html(titulo, valor, color_borde):
    return f"""
    <div style="background-color: white; border-top: 5px solid {color_borde}; padding: 20px; border-radius: 5px; 
    box-shadow: 0 2px 4px rgba(0,0,0,0.1); text-align: center; margin-bottom: 20px;">
        <p style="color: #666; font-size: 14px; margin: 0; font-family: Arial;">{titulo}</p>
        <p style="color: {COLOR_TEXTO_TITULO}; font-size: 28px; font-weight: bold; margin: 10px 0 0 0; font-family: Arial;">{valor}</p>
    </div>
    """

# --- CLASE PDF PROFESIONAL ---
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

    def footer(self):
        self.set_y(-15)
        self.set_font("Arial", "I", 8)
        self.set_text_color(128, 128, 128)
        self.cell(0, 10, str(self.page_no()), 0, 0, "C")

    def draw_table(self, title, df):
        if df.empty: return
        if df.index.name is not None: df = df.reset_index()
        
        # Evitar que el título quede solo al final de la hoja
        if self.get_y() + 30 > self.h - 20: self.add_page()

        self.set_font("Arial", "B", 12)
        self.set_text_color(*COLOR_TEXTO_TITULO)
        self.cell(0, 10, title, ln=True)
        
        def dibujar_encabezados(anchos, columnas):
            self.set_font("Arial", "B", 8)
            self.set_fill_color(70, 130, 180)
            self.set_text_color(255, 255, 255)
            for col in columnas:
                self.cell(anchos, 8, str(col), 1, 0, "C", True)
            self.ln()

        col_widths = self.page_width / len(df.columns)
        dibujar_encabezados(col_widths, df.columns)
        
        self.set_font("Arial", "", 8)
        self.set_text_color(50, 50, 50)
        for i, row in df.reset_index(drop=True).iterrows():
            # Repetir encabezados en página nueva
            if self.get_y() + 10 > self.h - 20:
                self.add_page()
                dibujar_encabezados(col_widths, df.columns)
                self.set_font("Arial", "", 8)
                self.set_text_color(50, 50, 50)
            
            fill = (i % 2 == 1)
            self.set_fill_color(240, 242, 246)
            if "TOTAL" in str(row.iloc[0]).upper():
                self.set_font("Arial", "B", 8)
                fill = False
            else:
                self.set_font("Arial", "", 8)
            
            for val in row:
                self.cell(col_widths, 7, str(val), 1, 0, "C", fill)
            self.ln()
        self.ln(5)

# --- PROCESO DE DATOS ---
def procesar_flujo_rrhh(archivo, f_inicio, f_fin):
    df_base = pd.read_excel(archivo, sheet_name='BaseQuery')
    df_activos_viejos = pd.read_excel(archivo, sheet_name='Activos')
    mapping = {'Gr.prof.': 'Categoría', 'División de personal': 'Línea', 'Division de personal': 'Línea'}
    df_base.rename(columns=mapping, inplace=True); df_activos_viejos.rename(columns=mapping, inplace=True)
    
    try:
        df_co_manual = pd.read_excel(archivo, sheet_name='CO')
        df_co_manual.rename(columns=mapping, inplace=True)
    except:
        df_co_manual = pd.DataFrame(columns=['Nº pers.', 'Apellido', 'Nombre de pila', 'Línea', 'Categoría', 'Desde', 'Motivo'])

    for df in [df_base, df_activos_viejos, df_co_manual]:
        if 'Nº pers.' in df.columns: df['Nº pers.'] = df['Nº pers.'].astype(str).str.strip()

    # Bajas (Lógica: Desde - 1 día para fecha real)
    df_bajas_raw = df_base[df_base['Status ocupación'] == 'Dado de baja'].copy()
    df_bajas_raw['Desde'] = pd.to_datetime(df_bajas_raw['Desde'])
    df_bajas_raw['Fecha_Real'] = df_bajas_raw['Desde'] - pd.Timedelta(days=1)
    df_bajas = df_bajas_raw[(df_bajas_raw['Fecha_Real'] >= pd.to_datetime(f_inicio)) & (df_bajas_raw['Fecha_Real'] <= pd.to_datetime(f_fin))].copy()
    df_bajas['Tipo'] = 'Baja'

    # Cambios Organizativos (Comparación de legajos)
    ids_desaparecidos = set(df_activos_viejos['Nº pers.']) - set(df_base['Nº pers.'])
    df_co_detectados = df_co_manual[df_co_manual['Nº pers.'].isin(ids_desaparecidos)].copy()
    if not df_co_detectados.empty:
        df_co_detectados['Fecha_Real'] = pd.to_datetime(df_co_detectados['Desde'])
        df_co_detectados['Tipo'] = 'Cambio Organizativo'
        df_co_detectados['Motivo de la medida'] = df_co_detectados['Motivo'] if 'Motivo' in df_co_detectados.columns else 'Reubicado'
        df_co = df_co_detectados[(df_co_detectados['Fecha_Real'] >= pd.to_datetime(f_inicio)) & (df_co_detectados['Fecha_Real'] <= pd.to_datetime(f_fin))].copy()
    else:
        df_co = pd.DataFrame(columns=['Nº pers.', 'Apellido', 'Nombre de pila', 'Línea', 'Categoría', 'Fecha_Real', 'Motivo de la medida', 'Tipo'])

    columnas = ['Nº pers.', 'Apellido', 'Nombre de pila', 'Línea', 'Categoría', 'Fecha_Real', 'Motivo de la medida', 'Tipo']
    df_final = pd.concat([df_bajas.reindex(columns=columnas), df_co.reindex(columns=columnas)], ignore_index=True)
    return df_final.sort_values('Fecha_Real'), len(df_base[df_base['Status ocupación'] == 'Activo'])

# --- APLICACIÓN ---
st.set_page_config(page_title="Gestión de Bajas y CO", layout="wide")
st.title("📊 Análisis de Salidas: Bajas y Cambios Org.")

st.sidebar.header("Configuración")
f_inicio = st.sidebar.date_input("Inicio", datetime(2025, 6, 1))
f_fin = st.sidebar.date_input("Fin", datetime(2026, 1, 15))

archivo = st.file_uploader("Subir Excel con pestañas BaseQuery, Activos, CO", type=['xlsx'])

if archivo:
    try:
        df_salidas, total_activos = procesar_flujo_rrhh(archivo, f_inicio, f_fin)
        if not df_salidas.empty:
            k1, k2, k3 = st.columns(3)
            k1.markdown(estilo_kpi_html("Dotación Activa", f"{total_activos:,}".replace(',', '.'), "#0476D0"), unsafe_allow_html=True)
            k2.markdown(estilo_kpi_html("Bajas del Período", len(df_salidas[df_salidas['Tipo'] == 'Baja']), "#EF553B"), unsafe_allow_html=True)
            k3.markdown(estilo_kpi_html("Cambio Organizativo", len(df_salidas[df_salidas['Tipo'] == 'Cambio Organizativo']), "#FFA500"), unsafe_allow_html=True)

            resumen_motivos = df_salidas.groupby(['Motivo de la medida', 'Tipo']).size().unstack(fill_value=0)
            if 'Baja' not in resumen_motivos.columns: resumen_motivos['Baja'] = 0
            if 'Cambio Organizativo' not in resumen_motivos.columns: resumen_motivos['Cambio Organizativo'] = 0
            resumen_motivos['Total'] = resumen_motivos.sum(axis=1)
            resumen_motivos.loc['TOTAL GENERAL'] = resumen_motivos.sum()
            st.write("### 📝 Motivos de Salida")
            st.dataframe(resumen_motivos.sort_values('Total', ascending=False), use_container_width=True)

            # Gráfico con Meses en Español
            df_salidas['Mes_Num'] = df_salidas['Fecha_Real'].dt.month
            df_salidas['Anio'] = df_salidas['Fecha_Real'].dt.year
            df_salidas['Mes_Display'] = df_salidas['Mes_Num'].map(MESES_ES) + " " + df_salidas['Anio'].astype(str)
            df_salidas['Mes_Sort'] = df_salidas['Fecha_Real'].dt.strftime('%Y-%m')
            
            df_grafico = df_salidas.groupby(['Mes_Sort', 'Mes_Display', 'Tipo']).size().reset_index(name='Cantidad').sort_values('Mes_Sort')

            fig = px.bar(df_grafico, x='Mes_Display', y='Cantidad', color='Tipo', barmode='group', text='Cantidad',
                         color_discrete_map={'Baja': '#EF553B', 'Cambio Organizativo': '#FFA500'})
            fig.update_traces(textposition='outside')
            st.plotly_chart(fig, use_container_width=True)

            st.write("### 👥 Detalle de Personas")
            st.dataframe(df_salidas[['Nº pers.', 'Apellido', 'Nombre de pila', 'Línea', 'Categoría', 'Fecha_Real', 'Motivo de la medida', 'Tipo']], hide_index=True)

            if st.button("📄 Generar Reporte PDF"):
                # SOLUCIÓN AL ERROR BytesIO: Usar archivo temporal
                img_bytes = fig.to_image(format="png", width=1000, height=500)
                with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
                    tmp.write(img_bytes)
                    tmp_path = tmp.name

                pdf = PDF(orientation='L', unit='mm', format='A4')
                pdf.report_title = f"Reporte de Bajas y C.O. ({f_inicio.strftime('%d/%m/%Y')} a {f_fin.strftime('%d/%m/%Y')})"
                pdf.add_page()
                
                # Insertar imagen desde el archivo temporal
                pdf.set_font("Arial", "B", 12)
                pdf.cell(0, 10, "Evolución Mensual", ln=True)
                pdf.image(tmp_path, x=10, y=None, w=220)
                pdf.ln(5)
                
                pdf.draw_table("Resumen de Motivos", resumen_motivos)
                pdf_rep = df_salidas[['Nº pers.', 'Apellido', 'Línea', 'Fecha_Real', 'Motivo de la medida', 'Tipo']].copy().rename(columns={'Fecha_Real': 'Fecha Real'})
                pdf.draw_table("Detalle de Bajas y C.O.", pdf_rep.astype(str))
                
                # Limpiar archivo temporal después de generar el PDF
                pdf_bytes = pdf.output(dest='S').encode('latin-1', 'replace')
                os.unlink(tmp_path)
                
                st.download_button("Descargar PDF", data=pdf_bytes, file_name=f"Reporte_RRHH_{f_fin}.pdf", mime="application/pdf")
        else: st.warning("Sin datos para el período.")
    except Exception as e: st.error(f"Error: {e}")
