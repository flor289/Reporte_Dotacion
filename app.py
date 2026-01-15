import streamlit as st
import pandas as pd
import plotly.express as px
from fpdf import FPDF
from datetime import datetime, timedelta
import io

# --- CONFIGURACIÓN Y ESTILOS ---
COLOR_AZUL_INSTITUCIONAL = (4, 118, 208)
COLOR_NARANJA_CO = (255, 165, 0)
COLOR_ROJO_BAJA = (239, 85, 59)
COLOR_TEXTO_TITULO = (0, 51, 102)

# Estilo CSS para los "Globos" (KPIs)
def estilo_kpi_html(titulo, valor, color_borde):
    return f"""
    <div style="
        background-color: white;
        border-top: 5px solid {color_borde};
        padding: 20px;
        border-radius: 5px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        text-align: center;
        margin-bottom: 20px;
    ">
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
        # Número de página en el centro inferior
        self.set_y(-15)
        self.set_font("Arial", "I", 8)
        self.set_text_color(128, 128, 128)
        self.cell(0, 10, str(self.page_no()), 0, 0, "C")

    def draw_table(self, title, df):
        if df.empty: return
        
        # Preparar datos (pasar índice a columna si tiene nombre)
        if df.index.name is not None:
            df = df.reset_index()
        
        # Verificar espacio para título y cabecera (evitar que el título quede solo al final)
        if self.get_y() + 30 > self.h - 20:
            self.add_page()

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
            # Salto de página automático con repetición de cabecera
            if self.get_y() + 10 > self.h - 20:
                self.add_page()
                dibujar_encabezados(col_widths, df.columns)
                self.set_font("Arial", "", 8)
                self.set_text_color(50, 50, 50)

            fill = (i % 2 == 1)
            self.set_fill_color(240, 242, 246)
            
            # Resaltar fila de TOTAL
            if "TOTAL" in str(row.iloc[0]).upper():
                self.set_font("Arial", "B", 8)
                fill = False
            else:
                self.set_font("Arial", "", 8)
            
            for val in row:
                self.cell(col_widths, 7, str(val), 1, 0, "C", fill)
            self.ln()
        self.ln(5)

# --- FUNCIONES DE PROCESO ---

def procesar_flujo_rrhh(archivo, f_inicio, f_fin):
    df_base = pd.read_excel(archivo, sheet_name='BaseQuery')
    df_activos_viejos = pd.read_excel(archivo, sheet_name='Activos')
    
    mapping = {'Gr.prof.': 'Categoría', 'División de personal': 'Línea', 'Division de personal': 'Línea'}
    df_base.rename(columns=mapping, inplace=True)
    df_activos_viejos.rename(columns=mapping, inplace=True)

    try:
        df_co_manual = pd.read_excel(archivo, sheet_name='CO')
        df_co_manual.rename(columns=mapping, inplace=True)
    except:
        df_co_manual = pd.DataFrame(columns=['Nº pers.', 'Apellido', 'Nombre de pila', 'Línea', 'Categoría', 'Desde', 'Motivo'])

    for df in [df_base, df_activos_viejos, df_co_manual]:
        if 'Nº pers.' in df.columns: df['Nº pers.'] = df['Nº pers.'].astype(str).str.strip()

    # 1. Bajas Sistema (Filtrado por fecha Desde - 1 día)
    df_bajas_raw = df_base[df_base['Status ocupación'] == 'Dado de baja'].copy()
    df_bajas_raw['Desde'] = pd.to_datetime(df_bajas_raw['Desde'])
    df_bajas_raw['Fecha_Real'] = df_bajas_raw['Desde'] - pd.Timedelta(days=1)
    mask_bajas = (df_bajas_raw['Fecha_Real'] >= pd.to_datetime(f_inicio)) & (df_bajas_raw['Fecha_Real'] <= pd.to_datetime(f_fin))
    df_bajas = df_bajas_raw[mask_bajas].copy()
    df_bajas['Tipo'] = 'Baja'

    # 2. C.O. (Por comparación de desaparición)
    ids_desaparecidos = set(df_activos_viejos['Nº pers.']) - set(df_base['Nº pers.'])
    df_co_detectados = df_co_manual[df_co_manual['Nº pers.'].isin(ids_desaparecidos)].copy()
    if not df_co_detectados.empty:
        df_co_detectados['Fecha_Real'] = pd.to_datetime(df_co_detectados['Desde'])
        df_co_detectados['Tipo'] = 'Cambio Organizativo'
        df_co_detectados['Motivo de la medida'] = df_co_detectados['Motivo'] if 'Motivo' in df_co_detectados.columns else 'Reubicado'
        mask_co = (df_co_detectados['Fecha_Real'] >= pd.to_datetime(f_inicio)) & (df_co_detectados['Fecha_Real'] <= pd.to_datetime(f_fin))
        df_co = df_co_detectados[mask_co].copy()
    else:
        df_co = pd.DataFrame(columns=['Nº pers.', 'Apellido', 'Nombre de pila', 'Línea', 'Categoría', 'Fecha_Real', 'Motivo de la medida', 'Tipo'])

    columnas = ['Nº pers.', 'Apellido', 'Nombre de pila', 'Línea', 'Categoría', 'Fecha_Real', 'Motivo de la medida', 'Tipo']
    df_final = pd.concat([df_bajas.reindex(columns=columnas), df_co.reindex(columns=columnas)], ignore_index=True)
    
    total_activos = len(df_base[df_base['Status ocupación'] == 'Activo'])

    return df_final.sort_values('Fecha_Real'), total_activos

# --- INTERFAZ STREAMLIT ---
st.set_page_config(page_title="Gestión de Bajas y CO", layout="wide")
st.title("📊 Análisis de Salidas: Bajas y Cambios Org.")

st.sidebar.header("Configuración del Reporte")
f_inicio = st.sidebar.date_input("Fecha Inicio", datetime(2025, 6, 1))
f_fin = st.sidebar.date_input("Fecha Fin", datetime(2026, 1, 15))

archivo = st.file_uploader("Subir Excel (BaseQuery, Activos, CO)", type=['xlsx'])

if archivo:
    try:
        df_salidas, total_activos = procesar_flujo_rrhh(archivo, f_inicio, f_fin)
        
        if not df_salidas.empty:
            st.subheader(f"Indicadores del Período: {f_inicio.strftime('%d/%m/%Y')} - {f_fin.strftime('%d/%m/%Y')}")
            
            # GLOBOS KPIs
            k1, k2, k3 = st.columns(3)
            val_activos = f"{total_activos:,}".replace(',', '.')
            k1.markdown(estilo_kpi_html("Dotación Activa", val_activos, "#0476D0"), unsafe_allow_html=True)
            
            bajas_n = len(df_salidas[df_salidas['Tipo'] == 'Baja'])
            k2.markdown(estilo_kpi_html("Bajas del Período", bajas_n, "#EF553B"), unsafe_allow_html=True)
            
            co_n = len(df_salidas[df_salidas['Tipo'] == 'Cambio Organizativo'])
            k3.markdown(estilo_kpi_html("Cambio Organizativo", co_n, "#FFA500"), unsafe_allow_html=True)

            # Cuadro de Motivos con Totales
            st.write("### 📝 Motivos de Salida (Bajas + CO)")
            resumen_motivos = df_salidas.groupby(['Motivo de la medida', 'Tipo']).size().unstack(fill_value=0)
            if 'Baja' not in resumen_motivos.columns: resumen_motivos['Baja'] = 0
            if 'Cambio Organizativo' not in resumen_motivos.columns: resumen_motivos['Cambio Organizativo'] = 0
            resumen_motivos['Total'] = resumen_motivos.sum(axis=1)
            resumen_motivos = resumen_motivos.sort_values('Total', ascending=False)
            resumen_motivos.loc['TOTAL GENERAL'] = resumen_motivos.sum()
            st.dataframe(resumen_motivos, use_container_width=True)

            # Gráfico de Barras con Cantidades
            st.write("### 📈 Evolución Mensual")
            df_salidas['Mes'] = df_salidas['Fecha_Real'].dt.strftime('%Y-%m')
            df_grafico = df_salidas.groupby(['Mes', 'Tipo']).size().reset_index(name='Cantidad')
            
            fig = px.bar(df_grafico, x='Mes', y='Cantidad', color='Tipo', barmode='group',
                         text='Cantidad',
                         color_discrete_map={'Baja': '#EF553B', 'Cambio Organizativo': '#FFA500'})
            fig.update_traces(textposition='outside')
            st.plotly_chart(fig, use_container_width=True)

            # Detalle Nominal
            st.write("### 👥 Detalle de Personas")
            st.dataframe(df_salidas.sort_values('Fecha_Real'), hide_index=True)

            if st.button("📄 Generar Reporte PDF"):
                pdf = PDF(orientation='L', unit='mm', format='A4')
                pdf.report_title = f"Reporte de Bajas y C.O. ({f_inicio.strftime('%d/%m/%Y')} a {f_fin.strftime('%d/%m/%Y')})"
                pdf.add_page()
                
                # Tabla 1: Motivos
                pdf.draw_table("Resumen de Motivos", resumen_motivos)
                
                # Tabla 2: Detalle con títulos repetidos
                columnas_rep = ['Nº pers.', 'Apellido', 'Línea', 'Fecha_Real', 'Motivo de la medida', 'Tipo']
                df_rep = df_salidas[columnas_rep].copy()
                df_rep.rename(columns={'Fecha_Real': 'Fecha Real'}, inplace=True)
                
                pdf.draw_table("Detalle de Bajas y C.O.", df_rep.astype(str))
                
                pdf_output = pdf.output(dest='S').encode('latin-1', 'replace')
                st.download_button("Descargar PDF", data=pdf_output, file_name=f"Reporte_Salidas_{f_fin}.pdf", mime="application/pdf")
        else:
            st.warning("No se detectaron salidas para el rango seleccionado.")
    except Exception as e:
        st.error(f"Error crítico: {e}")
