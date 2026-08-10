import streamlit as st
import pandas as pd
from fpdf import FPDF
from datetime import datetime, timedelta
import io

# --- 1. CONFIGURACIÓN Y ESTILOS ---
COLOR_AZUL_INSTITUCIONAL = (4, 118, 208)
COLOR_FONDO_CABECERA_TABLA = (70, 130, 180)
COLOR_GRIS_FONDO_FILA = (240, 242, 246)
COLOR_GRIS_LINEA = (220, 220, 220)
COLOR_TEXTO_TITULO = (0, 51, 102)
COLOR_TEXTO_CUERPO = (50, 50, 50)
COLOR_CELESTE_PASTEL = (186, 225, 255)  # Celeste pastel para Cambio Categoría
COLOR_AZUL_PASTEL_OSCURO = (120, 180, 235)  # Celeste/Azul un poco más oscuro para Cambio Línea

class PDF(FPDF):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.page_width = self.w - 2 * self.l_margin
        self.report_title = "Resumen de Dotación"

    def header(self):
        self.set_font("Arial", "B", 18)
        self.set_text_color(*COLOR_TEXTO_TITULO)
        self.cell(0, 10, self.report_title, 0, 0, "C")
        self.ln(15)

    def footer(self):
        self.set_y(-15)
        self.set_font("Arial", "I", 8)
        self.set_text_color(128, 128, 128)
        self.cell(0, 10, str(self.page_no()), 0, 0, "C")

    def draw_section_title(self, title):
        self.set_font("Arial", "B", 14)
        self.set_text_color(*COLOR_TEXTO_TITULO)
        self.cell(0, 10, title, ln=True, align="L")
        self.set_draw_color(*COLOR_AZUL_INSTITUCIONAL)
        self.set_line_width(0.5)
        self.line(self.get_x(), self.get_y(), self.get_x() + self.page_width, self.get_y())
        self.ln(5)

    def draw_kpi_box(self, title, value, color, x, y, width=80):
        kpi_height = 16
        self.set_xy(x, y)
        self.set_fill_color(*color)
        self.cell(width, 1.5, "", fill=True, ln=False, border=0)
        self.set_xy(x, y + 1.5)
        self.set_fill_color(255, 255, 255)
        self.set_draw_color(*COLOR_GRIS_LINEA)
        self.cell(width, kpi_height - 1.5, "", border=1, fill=True)
        self.set_xy(x, y + 3)
        self.set_font('Arial', '', 10)
        self.set_text_color(*COLOR_TEXTO_CUERPO)
        self.cell(width, 8, title, align='C')
        self.set_xy(x, y + 8)
        self.set_font('Arial', 'B', 16)
        self.set_text_color(*COLOR_TEXTO_TITULO)
        self.cell(width, 10, str(value), align='C')

    def draw_table(self, title, df_original, is_crosstab=False):
        if df_original.empty: return
        df = df_original.copy()
        if is_crosstab: 
            df = df.replace(0, '-')
            if df.index.name: df.reset_index(inplace=True)
        
        if self.get_y() + (8 * (len(df) + 1) + 10) > self.h - self.b_margin: self.add_page(orientation=self.cur_orientation)
        self.draw_section_title(title)
        
        df_formatted = df.copy()
        for col in df_formatted.columns:
            if pd.api.types.is_numeric_dtype(df_formatted[col]) and col not in ['Nº pers.', 'Antigüedad', 'Edad']:
                if "Prom." in str(col):
                    df_formatted[col] = df_formatted[col].apply(lambda x: f"{round(x):.0f}" if isinstance(x, (int, float)) else x)
                else:
                    df_formatted[col] = df_formatted[col].apply(lambda x: f"{x:,.0f}".replace(',', '.') if isinstance(x, (int, float)) else x)
        
        widths = {col: max(self.get_string_width(str(col)) + 10, df_formatted[col].astype(str).apply(lambda x: self.get_string_width(x)).max() + 10) for col in df_formatted.columns}
        total_width = sum(widths.values())
        if total_width > self.page_width:
            scaling_factor = self.page_width / total_width
            widths = {k: v * scaling_factor for k, v in widths.items()}
        
        self.set_font("Arial", "B", 8)
        self.set_fill_color(*COLOR_FONDO_CABECERA_TABLA)
        self.set_text_color(255, 255, 255)
        for col in df_formatted.columns:
            self.cell(widths[col], 8, str(col), 0, 0, "C", True)
        self.ln()
        
        self.set_text_color(*COLOR_TEXTO_CUERPO)
        self.set_draw_color(*COLOR_GRIS_LINEA)
        self.set_line_width(0.2)
        
        for i, (_, row) in enumerate(df_formatted.iterrows()):
            if self.get_y() + 8 > self.h - self.b_margin:
                self.add_page(orientation=self.cur_orientation)
                self.set_font("Arial", "B", 8)
                self.set_fill_color(*COLOR_FONDO_CABECERA_TABLA)
                self.set_text_color(255, 255, 255)
                for col in df_formatted.columns:
                    self.cell(widths[col], 8, str(col), 0, 0, "C", True)
                self.ln()
                self.set_text_color(*COLOR_TEXTO_CUERPO)
            
            fill = i % 2 == 1
            self.set_font("Arial", "B" if "Total" in str(row.iloc[0]) else "", 8)
            self.set_fill_color(*COLOR_GRIS_FONDO_FILA)
            for col in df_formatted.columns:
                self.cell(widths[col], 8, str(row[col]), 'T', 0, "C", fill)
            self.ln()
        self.ln(10)

# --- 2. LÓGICA DE CÁLCULO ---
def calcular_años(fecha_inicio, fecha_fin):
    if pd.isna(fecha_inicio) or pd.isna(fecha_fin): return 0
    return (fecha_fin - fecha_inicio).days / 365.25

def generar_resumen_completo(df_datos, index_col='Categoría', columns_col='Línea', incluir_promedios=True):
    if df_datos.empty: return pd.DataFrame()
    resumen = pd.crosstab(df_datos[index_col], df_datos[columns_col], margins=True, margins_name="Total")
    
    if incluir_promedios and 'Antigüedad' in df_datos.columns and 'Edad' in df_datos.columns:
        promedios = df_datos.groupby(index_col).agg({'Antigüedad': 'mean', 'Edad': 'mean'})
        promedios.loc['Total', 'Antigüedad'] = df_datos['Antigüedad'].mean()
        promedios.loc['Total', 'Edad'] = df_datos['Edad'].mean()
        resumen['Antig. Prom.'] = promedios['Antigüedad']
        resumen['Edad Prom.'] = promedios['Edad']
        
    return resumen

def procesar_recategorizaciones(df_base, df_activos_prev):
    """Detecta quiénes están activos hoy y cambiaron su categoría respecto a la foto de Activos."""
    if df_base.empty or df_activos_prev.empty:
        return pd.DataFrame()
    
    df_act_hoy = df_base[df_base['Status ocupación'] == 'Activo'].copy()
    df_act_viejos = df_activos_prev.copy()

    mapping = {'Gr.prof.': 'Categoría', 'División de personal': 'Línea', 'Division de personal': 'Línea'}
    df_act_hoy.rename(columns=mapping, inplace=True)
    df_act_viejos.rename(columns=mapping, inplace=True)

    if 'Nº pers.' in df_act_hoy.columns: df_act_hoy['Nº pers.'] = df_act_hoy['Nº pers.'].astype(str).str.strip()
    if 'Nº pers.' in df_act_viejos.columns: df_act_viejos['Nº pers.'] = df_act_viejos['Nº pers.'].astype(str).str.strip()

    if 'Categoría' in df_act_hoy.columns and 'Categoría' in df_act_viejos.columns:
        df_cmp = pd.merge(
            df_act_hoy[['Nº pers.', 'Apellido', 'Nombre de pila', 'Línea', 'Categoría']],
            df_act_viejos[['Nº pers.', 'Categoría']],
            on='Nº pers.',
            suffixes=('_Actual', '_Anterior'),
            how='inner'
        )
        df_recat = df_cmp[df_cmp['Categoría_Actual'] != df_cmp['Categoría_Anterior']].copy()
        df_recat.rename(columns={'Categoría_Anterior': 'Categoría Anterior', 'Categoría_Actual': 'Categoría Actual'}, inplace=True)
        return df_recat
    return pd.DataFrame()

def procesar_cambios_linea(df_base, df_activos_prev):
    """Detecta quiénes están activos hoy y cambiaron su Línea respecto a la foto de Activos."""
    if df_base.empty or df_activos_prev.empty:
        return pd.DataFrame()
    
    df_act_hoy = df_base[df_base['Status ocupación'] == 'Activo'].copy()
    df_act_viejos = df_activos_prev.copy()

    mapping = {'Gr.prof.': 'Categoría', 'División de personal': 'Línea', 'Division de personal': 'Línea'}
    df_act_hoy.rename(columns=mapping, inplace=True)
    df_act_viejos.rename(columns=mapping, inplace=True)

    if 'Nº pers.' in df_act_hoy.columns: df_act_hoy['Nº pers.'] = df_act_hoy['Nº pers.'].astype(str).str.strip()
    if 'Nº pers.' in df_act_viejos.columns: df_act_viejos['Nº pers.'] = df_act_viejos['Nº pers.'].astype(str).str.strip()

    if 'Línea' in df_act_hoy.columns and 'Línea' in df_act_viejos.columns:
        df_cmp = pd.merge(
            df_act_hoy[['Nº pers.', 'Apellido', 'Nombre de pila', 'Categoría', 'Línea']],
            df_act_viejos[['Nº pers.', 'Línea']],
            on='Nº pers.',
            suffixes=('_Actual', '_Anterior'),
            how='inner'
        )
        df_cambio_l = df_cmp[df_cmp['Línea_Actual'] != df_cmp['Línea_Anterior']].copy()
        df_cambio_l.rename(columns={'Línea_Anterior': 'Línea Anterior', 'Línea_Actual': 'Línea Actual'}, inplace=True)
        return df_cambio_l
    return pd.DataFrame()

# --- 3. PROCESAMIENTO ---
def procesar_archivo_base(archivo_cargado, sheet_name='BaseQuery'):
    try:
        df = pd.read_excel(archivo_cargado, sheet_name=sheet_name, engine='openpyxl')
        df.rename(columns={'Gr.prof.': 'Categoría', 'División de personal': 'Línea', 'Motivo de la medida': 'Motivo de Baja'}, inplace=True)
        for col in ['Fecha', 'Desde', 'Fecha nac.']:
            if col in df.columns: df[col] = pd.to_datetime(df[col], errors='coerce')
        orden_lineas = ['ROCA', 'MITRE', 'SARMIENTO', 'SAN MARTIN', 'BELGRANO SUR', 'REGIONALES', 'CENTRAL']
        orden_categorias = ['COOR.E.T', 'INST.TEC', 'INS.CERT', 'CON.ELEC', 'CON.DIES', 'AY.CON.H', 'AY.CONDU', 'ASP.AY.C']
        if 'Línea' in df.columns: df['Línea'] = pd.Categorical(df['Línea'], categories=orden_lineas, ordered=True)
        if 'Categoría' in df.columns: df['Categoría'] = pd.Categorical(df['Categoría'], categories=orden_categorias, ordered=True)
        return df
    except: return pd.DataFrame()

def procesar_metricas_novedades(df_altas_raw, df_bajas_raw, df_co_raw, fecha_ref):
    df_bajas = df_bajas_raw.copy()
    if not df_bajas.empty:
        df_bajas['Antigüedad'] = df_bajas.apply(lambda r: calcular_años(r['Fecha'], r['Desde']), axis=1)
        df_bajas['Edad'] = df_bajas.apply(lambda r: calcular_años(r['Fecha nac.'], r['Desde']), axis=1)
        df_bajas_vis = df_bajas.copy()
        df_bajas_vis['Antigüedad'] = df_bajas_vis['Antigüedad'].apply(lambda x: int(round(x)))
        df_bajas_vis['Fecha nac.'] = df_bajas_vis['Fecha nac.'].dt.strftime('%d/%m/%Y')
        df_bajas_vis['Desde'] = df_bajas_vis['Desde'].dt.strftime('%d/%m/%Y')
    else: df_bajas_vis = pd.DataFrame()

    df_altas = df_altas_raw.copy()
    if not df_altas.empty:
        df_altas['Antigüedad'] = df_altas.apply(lambda r: calcular_años(r['Fecha'], fecha_ref), axis=1)
        df_altas['Edad'] = df_altas.apply(lambda r: calcular_años(r['Fecha nac.'], fecha_ref), axis=1)
        df_altas_vis = df_altas.copy()
        df_altas_vis['Fecha'] = df_altas_vis['Fecha'].dt.strftime('%d/%m/%Y')
        df_altas_vis['Fecha nac.'] = df_altas_vis['Fecha nac.'].dt.strftime('%d/%m/%Y')
    else: df_altas_vis = pd.DataFrame()

    df_co = df_co_raw.copy()
    if not df_co.empty and 'Desde' in df_co.columns:
        df_co['Antigüedad'] = df_co.apply(lambda r: calcular_años(r['Fecha'], r['Desde']), axis=1)
        df_co['Edad'] = df_co.apply(lambda r: calcular_años(r['Fecha nac.'], r['Desde']), axis=1)
        df_co_vis = df_co.copy()
        if pd.api.types.is_datetime64_any_dtype(df_co_vis['Desde']):
            df_co_vis['Desde'] = df_co_vis['Desde'].dt.strftime('%d/%m/%Y')
    else: df_co_vis = pd.DataFrame()

    return df_altas, df_altas_vis, df_bajas, df_bajas_vis, df_co, df_co_vis

def filtrar_novedades_por_fecha(df_base_para_filtrar, fecha_inicio, fecha_fin):
    df = df_base_para_filtrar.copy()
    altas_filtradas = df[(df['Fecha'] >= fecha_inicio) & (df['Fecha'] <= fecha_fin)].copy()
    df_bajas_p = df[df['Status ocupación'] == 'Dado de baja'].copy()
    if not df_bajas_p.empty:
        df_bajas_p['f_corregida'] = df_bajas_p['Desde'] - pd.Timedelta(days=1)
        bajas_f = df_bajas_p[(df_bajas_p['f_corregida'] >= fecha_inicio) & (df_bajas_p['f_corregida'] <= fecha_fin)].copy()
        if not bajas_f.empty: bajas_f['Desde'] = bajas_f['f_corregida']
    else: bajas_f = pd.DataFrame()
    return altas_filtradas, bajas_f

def crear_pdf_reporte(titulo_reporte, rango_fechas_str, df_altas, df_bajas, res_altas, res_bajas, res_activos, res_bajas_linea, res_bajas_cat, df_co=None, df_recat=None, df_cambio_linea=None):
    pdf = PDF(orientation='L', unit='mm', format='A4')
    pdf.report_title = titulo_reporte
    pdf.add_page()
    pdf.draw_section_title(f"Indicadores del Período: {rango_fechas_str}")
    total_act = f"{res_activos.loc['Total', 'Total']:,}".replace(',', '.') if not res_activos.empty else "0"
    
    has_co = df_co is not None and not df_co.empty
    has_recat = df_recat is not None and not df_recat.empty
    has_linea = df_cambio_linea is not None and not df_cambio_linea.empty

    num_kpis = 3 + (1 if has_co else 0) + (1 if has_recat else 0) + (1 if has_linea else 0)
    k_w = pdf.page_width / (num_kpis + 0.5)
    sp = (pdf.page_width - (k_w * num_kpis)) / max(1, (num_kpis - 1))
    
    y = pdf.get_y()
    curr_x = pdf.l_margin
    
    pdf.draw_kpi_box("Dotación Activa", total_act, (200, 200, 200), curr_x, y, width=k_w)
    curr_x += k_w + sp
    
    pdf.draw_kpi_box("Altas del Período", '-' if len(df_altas) == 0 else str(len(df_altas)), (200, 200, 200), curr_x, y, width=k_w)
    curr_x += k_w + sp
    
    pdf.draw_kpi_box("Bajas del Período", '-' if len(df_bajas) == 0 else str(len(df_bajas)), (200, 200, 200), curr_x, y, width=k_w)
    curr_x += k_w + sp
    
    if has_co:
        pdf.draw_kpi_box("Cambio Organizativo", str(len(df_co)), (255, 165, 0), curr_x, y, width=k_w)
        curr_x += k_w + sp
        
    if has_recat:
        pdf.draw_kpi_box("Cambio Categoría", str(len(df_recat)), COLOR_CELESTE_PASTEL, curr_x, y, width=k_w)
        curr_x += k_w + sp

    if has_linea:
        pdf.draw_kpi_box("Cambio Línea", str(len(df_cambio_linea)), COLOR_AZUL_PASTEL_OSCURO, curr_x, y, width=k_w)
    
    pdf.ln(22)
    pdf.draw_table(f"Composición de la Dotación Activa", res_activos, is_crosstab=True)
    pdf.draw_table(f"Resumen de Bajas (Período: {rango_fechas_str})", res_bajas, is_crosstab=True)
    pdf.draw_table("Motivos de Baja por Línea", res_bajas_linea, is_crosstab=True)
    pdf.draw_table("Motivos de Baja por Categoría", res_bajas_cat, is_crosstab=True)
    pdf.draw_table(f"Resumen de Altas (Período: {rango_fechas_str})", res_altas, is_crosstab=True)
    
    if not df_bajas.empty: pdf.draw_table("Detalle de Bajas", df_bajas[['Nº pers.', 'Apellido', 'Nombre de pila', 'Motivo de Baja', 'Fecha nac.', 'Antigüedad', 'Desde', 'Línea', 'Categoría']])
    if not df_altas.empty: pdf.draw_table("Detalle de Altas", df_altas[['Nº pers.', 'Apellido', 'Nombre de pila', 'Fecha nac.', 'Fecha', 'Línea', 'Categoría']])
    if has_co: pdf.draw_table("Detalle Cambios Organizativos", df_co[['Nº pers.', 'Apellido', 'Nombre de pila', 'Desde', 'Línea', 'Categoría']])
    if has_recat: pdf.draw_table("Detalle Cambios de Categoría", df_recat[['Nº pers.', 'Apellido', 'Nombre de pila', 'Línea', 'Categoría Anterior', 'Categoría Actual']])
    if has_linea: pdf.draw_table("Detalle Cambios de Línea", df_cambio_linea[['Nº pers.', 'Apellido', 'Nombre de pila', 'Categoría', 'Línea Anterior', 'Línea Actual']])
    
    return pdf.output(dest='S').encode('latin-1', 'replace')

# --- 4. INTERFAZ ---
st.set_page_config(page_title="Dashboard de Dotación", layout="wide")
st.title("📊 Dashboard de Control de Dotación")

tabs = st.tabs(["📅 Reporte Diario", "📅 Semanal", "📅 Mensual", "📅 Anual"])

with tabs[0]:
    st.header("Análisis Diario")
    uploaded_file = st.file_uploader("Sube tu archivo Excel", type=['xlsx'], key="up_main")
    if uploaded_file:
        try:
            df_base = procesar_archivo_base(uploaded_file, 'BaseQuery')
            df_act_p = pd.read_excel(uploaded_file, sheet_name='Activos')
            try: df_co_r = procesar_archivo_base(uploaded_file, 'CO')
            except: df_co_r = pd.DataFrame()

            v_legs = set(df_act_p['Nº pers.'].astype(str).str.strip()) if 'Nº pers.' in df_act_p.columns else set()
            desap = v_legs - set(df_base['Nº pers.'].astype(str).str.strip())
            df_co_raw = df_co_r[df_co_r['Nº pers.'].astype(str).str.strip().isin(desap)].copy() if not df_co_r.empty else df_act_p[df_act_p['Nº pers.'].astype(str).str.strip().isin(desap)].copy()

            df_alt_r = df_base[~df_base['Nº pers.'].astype(str).str.strip().isin(v_legs) & (df_base['Status ocupación'] == 'Activo')].copy()
            df_baj_r = df_base[df_base['Nº pers.'].astype(str).str.strip().isin(v_legs) & (df_base['Status ocupación'] == 'Dado de baja')].copy()
            
            if not df_baj_r.empty: 
                df_baj_r['Desde'] = df_baj_r['Desde'] - pd.Timedelta(days=1)
                df_baj_r = df_baj_r.sort_values(by='Desde', ascending=True)
            if not df_alt_r.empty: df_alt_r = df_alt_r.sort_values(by='Fecha', ascending=True)

            hoy = pd.to_datetime(datetime.now())
            df_a, df_a_v, df_b, df_b_v, df_c, df_c_v = procesar_metricas_novedades(df_alt_r, df_baj_r, df_co_raw, hoy)
            
            # PROCESO DE CAMBIOS DE CATEGORÍA Y LÍNEA (Exclusivo para PDF)
            df_recat = procesar_recategorizaciones(df_base, df_act_p)
            df_cambio_l = procesar_cambios_linea(df_base, df_act_p)

            df_act_h = df_base[df_base['Status ocupación'] == 'Activo'].copy()
            df_act_h['Antigüedad'] = df_act_h.apply(lambda r: calcular_años(r['Fecha'], hoy), axis=1)
            df_act_h['Edad'] = df_act_h.apply(lambda r: calcular_años(r['Fecha nac.'], hoy), axis=1)

            res_act = generar_resumen_completo(df_act_h)
            res_alt = generar_resumen_completo(df_a, incluir_promedios=False)
            res_baj = generar_resumen_completo(df_b)
            res_baj_linea = pd.crosstab(df_b['Motivo de Baja'], df_b['Línea'], margins=True, margins_name="Total") if not df_b.empty else pd.DataFrame()
            res_baj_cat = pd.crosstab(df_b['Motivo de Baja'], df_b['Categoría'], margins=True, margins_name="Total") if not df_b.empty else pd.DataFrame()

            pdf = crear_pdf_reporte("Resumen Diario de Dotación", datetime.now().strftime('%d/%m/%Y'), df_a_v, df_b_v, res_alt, res_baj, res_act, res_baj_linea, res_baj_cat, df_c_v, df_recat, df_cambio_l)
            st.download_button("📄 Descargar Reporte Diario", pdf, f"Reporte_Diario_Dotacion_{datetime.now().strftime('%Y%m%d')}.pdf", "application/pdf")

            st.session_state.uploaded_file = uploaded_file
            st.session_state.df_base = df_base
            st.session_state.df_activos_prev = df_act_p
            st.session_state.df_co_respaldo = df_co_r
        except Exception as e: st.error(f"Error: {e}")

def render_report(report_type):
    st.header(f"Generador de Reportes {report_type}es")
    if 'uploaded_file' in st.session_state:
        today = datetime.now()
        
        if report_type == 'Semanal': 
            d_s = today - timedelta(days=7); d_e = today
        elif report_type == 'Mensual': 
            d_s = today.replace(day=1); d_e = (d_s + timedelta(days=32)).replace(day=1) - timedelta(days=1)
        else: 
            d_s = today.replace(month=1, day=1); d_e = today.replace(month=12, day=31)

        c1, c2 = st.columns(2)
        start = c1.date_input("Inicio", d_s, key=f"s_{report_type}")
        end = pd.to_datetime(c2.date_input("Fin", d_e, key=f"e_{report_type}"))
        
        if start and end and start <= end.date():
            df_base = st.session_state.df_base
            df_act_p = st.session_state.df_activos_prev
            df_co_r = st.session_state.df_co_respaldo
            df_alt_raw, df_baj_raw = filtrar_novedades_por_fecha(df_base, pd.to_datetime(start), end)
            
            if not df_alt_raw.empty: df_alt_raw = df_alt_raw.sort_values(by='Fecha', ascending=True)
            if not df_baj_raw.empty: df_baj_raw = df_baj_raw.sort_values(by='Desde', ascending=True)

            if report_type == 'Anual' and not df_alt_raw.empty:
                df_alt_raw['Categoría'] = 'ASP.AY.C'
                st.info("💡 Normalización anual a ASP.AY.C aplicada.")

            desap = set(df_act_p['Nº pers.'].astype(str).str.strip()) - set(df_base['Nº pers.'].astype(str).str.strip())
            df_co_f = df_co_r[(df_co_r['Nº pers.'].astype(str).str.strip().isin(desap)) & (df_co_r['Desde'] >= pd.to_datetime(start)) & (df_co_r['Desde'] <= end)].copy() if not df_co_r.empty else pd.DataFrame()
            if not df_co_f.empty: df_co_f = df_co_f.sort_values(by='Desde', ascending=True)

            df_a, df_a_v, df_b, df_b_v, df_c, df_c_v = procesar_metricas_novedades(df_alt_raw, df_baj_raw, df_co_f, end)
            
            df_recat = procesar_recategorizaciones(df_base, df_act_p)
            df_cambio_l = procesar_cambios_linea(df_base, df_act_p)

            df_act_per = df_base[(df_base['Fecha'] <= end) & (df_base['Status ocupación'] == 'Activo')].copy()
            df_act_per['Antigüedad'] = df_act_per.apply(lambda r: calcular_años(r['Fecha'], end), axis=1)
            df_act_per['Edad'] = df_act_per.apply(lambda r: calcular_años(r['Fecha nac.'], end), axis=1)

            res_act = generar_resumen_completo(df_act_per)
            res_alt = generar_resumen_completo(df_a, incluir_promedios=False)
            res_baj = generar_resumen_completo(df_b)
            res_baj_linea = pd.crosstab(df_b['Motivo de Baja'], df_b['Línea'], margins=True, margins_name="Total") if not df_b.empty else pd.DataFrame()
            res_baj_cat = pd.crosstab(df_b['Motivo de Baja'], df_b['Categoría'], margins=True, margins_name="Total") if not df_b.empty else pd.DataFrame()

            titulo_pdf = f"Reporte {report_type} de Dotación"
            pdf = crear_pdf_reporte(titulo_pdf, f"{start.strftime('%d/%m/%Y')} - {end.strftime('%d/%m/%Y')}", df_a_v, df_b_v, res_alt, res_baj, res_act, res_baj_linea, res_baj_cat, df_c_v, df_recat, df_cambio_l)
            
            if report_type == 'Anual':
                n_file = f"Reporte_Anual_{start.strftime('%Y')}.pdf"
            else:
                n_file = f"Reporte_{report_type}_{start.strftime('%Y%m%d')}_a_{end.strftime('%Y%m%d')}.pdf"
                
            st.download_button(f"📄 Descargar {titulo_pdf}", pdf, n_file, "application/pdf")

    else: st.info("Sube un archivo primero.")

with tabs[1]: render_report('Semanal')
with tabs[2]: render_report('Mensual')
with tabs[3]: render_report('Anual')
