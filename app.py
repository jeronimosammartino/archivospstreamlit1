import io
from datetime import datetime
import streamlit as st
import pandas as pd
import openpyxl
import plotly.express as px
import plotly.graph_objects as go
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image as RLImage, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

st.set_page_config(page_title="Control de Métodos, Costos y HH", layout="wide")

st.title("📊 Control de Métodos: Horas Hombre, Costos y Rendimientos")
st.write("Tablero integral de control operacional con exportación ejecutiva a PowerPoint y PDF.")

uploaded_file = st.file_uploader("Sube la matriz del proyecto (.xlsx)", type=["xlsx", "xls"])

def leer_fabricost(file_bytes):
    try:
        df_costos = pd.read_excel(file_bytes, sheet_name="FABRICOST")
        col_recurso, col_tarifa = None, None
        for col in df_costos.columns:
            col_str = str(col).strip().upper()
            if "RECURSO" in col_str:
                col_recurso = col
            elif any(k in col_str for k in ["COSTO", "CIF", "TARIFA", "VALOR"]):
                col_tarifa = col
        if col_recurso and col_tarifa:
            df_costos = df_costos[[col_recurso, col_tarifa]].dropna()
            df_costos.columns = ["Recurso", "Tarifa"]
            df_costos["Recurso_Clean"] = df_costos["Recurso"].astype(str).str.strip().str.upper()
            df_costos["Tarifa"] = pd.to_numeric(df_costos["Tarifa"], errors="coerce").fillna(0)
            return df_costos
    except Exception:
        pass
    return pd.DataFrame(columns=["Recurso_Clean", "Tarifa"])

def leer_matriz_exacta(file_bytes, sheet_name):
    wb = openpyxl.load_workbook(file_bytes, data_only=True)
    ws = wb[sheet_name]
    
    for mr in list(ws.merged_cells.ranges):
        val = ws.cell(row=mr.min_row, column=mr.min_col).value
        ws.unmerge_cells(range_string=str(mr))
        for r in range(mr.min_row, mr.max_row + 1):
            for c in range(mr.min_col, mr.max_col + 1):
                ws.cell(row=r, column=c).value = val
                
    filas = list(ws.iter_rows(values_only=True))
    if not filas:
        return pd.DataFrame()
        
    header_idx = 0
    for idx, row in enumerate(filas[:10]):
        textos = [str(c).strip().upper() for c in row if c is not None]
        if any("CATEGORIA" in t or "GDF" in t or "PESO" in t for t in textos):
            header_idx = idx
            break
            
    headers = [str(c).strip().replace("\n", " ") if c is not None else f"Col_{i}" for i, c in enumerate(filas[header_idx])]
    df = pd.DataFrame(filas[header_idx + 1:], columns=headers).dropna(how="all").reset_index(drop=True)
    
    if "Categoria" in df.columns:
        df["Categoria"] = df["Categoria"].ffill()
    return df

def procesar_datos(df, df_tarifas):
    cols_hh = [c for c in df.columns if c.startswith("HH ") or c in ["Autycontrol", "MET"]]
    df_det = df[~df['GDF'].astype(str).str.contains("Total", case=False, na=False)].copy()
    
    def normalizar_cat(val):
        v = str(val).upper()
        if "OFERTA" in v:
            return "OFERTADAS"
        elif "HDR" in v:
            return "ESTIMADAS"
        elif "REAL" in v:
            return "REALES"
        return None

    df_det['Estado'] = df_det['Categoria'].apply(normalizar_cat)
    df_det = df_det.dropna(subset=['Estado'])
    
    df_long = df_det.melt(
        id_vars=['Estado', 'GDF', 'Peso'],
        value_vars=cols_hh,
        var_name='Recurso',
        value_name='Horas'
    )
    df_long['Horas'] = pd.to_numeric(df_long['Horas'], errors='coerce').fillna(0)
    df_long['Peso'] = pd.to_numeric(df_long['Peso'], errors='coerce').fillna(0)
    
    df_long['Recurso_Clean'] = df_long['Recurso'].astype(str).str.strip().str.upper()
    if not df_tarifas.empty:
        df_long = pd.merge(df_long, df_tarifas[['Recurso_Clean', 'Tarifa']], on='Recurso_Clean', how='left')
        df_long['Tarifa'] = df_long['Tarifa'].fillna(0)
        df_long['Costo'] = df_long['Horas'] * df_long['Tarifa']
    else:
        df_long['Costo'] = 0.0
        
    df_pesos = df_det[['Estado', 'GDF', 'Peso']].drop_duplicates()
    df_pesos['Peso'] = pd.to_numeric(df_pesos['Peso'], errors='coerce').fillna(0)
    
    return df_long, df_pesos

def formato_m_k(val):
    if abs(val) >= 1_000_000:
        return f"{val/1_000_000:.2f}M"
    elif abs(val) >= 1_000:
        return f"{val/1_000:.2f}K"
    return f"{val:.2f}"

def render_chart_image(plot_type, data_dict):
    fig, ax = plt.subplots(figsize=(7.5, 4.2), dpi=160)
    fig.patch.set_facecolor('#FFFFFF')
    ax.set_facecolor('#F8FAFC')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_color('#CBD5E1')
    ax.spines['bottom'].set_color('#CBD5E1')
    ax.grid(axis='y', linestyle='--', alpha=0.6, color='#E2E8F0')
    ax.set_axisbelow(True)

    if plot_type == "global_horas":
        estados = ["OFERTADAS", "ESTIMADAS", "REALES"]
        vals = [data_dict['h_ofe'], data_dict['h_est'], data_dict['h_real']]
        bars = ax.bar(estados, vals, color=["#F59E0B", "#06B6D4", "#6366F1"], width=0.52)
        for b in bars:
            h = b.get_height()
            ax.text(b.get_x() + b.get_width()/2., h + (max(vals)*0.015), f"{int(h):,}", ha='center', va='bottom', fontsize=9, fontweight='bold', color='#1E293B')
        ax.set_ylabel("Horas Hombre (HH)", fontsize=9, color='#475569')

    elif plot_type == "global_costos":
        estados = ["OFERTADOS", "ESTIMADOS", "REALES"]
        vals = [data_dict['c_ofe'], data_dict['c_est'], data_dict['c_real']]
        bars = ax.bar(estados, vals, color=["#F59E0B", "#06B6D4", "#6366F1"], width=0.52)
        for b in bars:
            h = b.get_height()
            ax.text(b.get_x() + b.get_width()/2., h + (max(vals)*0.015), formato_m_k(h), ha='center', va='bottom', fontsize=9, fontweight='bold', color='#1E293B')
        ax.set_ylabel("Costo Total ($)", fontsize=9, color='#475569')

    elif plot_type == "rendimientos":
        estados = ["OFERTADO", "ESTIMADO", "REAL"]
        vals = [data_dict['rend_ofe'], data_dict['rend_est'], data_dict['rend_real']]
        bars = ax.bar(estados, vals, color=["#F59E0B", "#06B6D4", "#6366F1"], width=0.52)
        for b in bars:
            h = b.get_height()
            ax.text(b.get_x() + b.get_width()/2., h + 0.05, f"{h:.2f}", ha='center', va='bottom', fontsize=9, fontweight='bold', color='#1E293B')
        ax.set_ylabel("Rendimiento (Kg / HH)", fontsize=9, color='#475569')

    elif plot_type == "top_desvios":
        desv = data_dict['desvios_tarea'].head(7)
        bars = ax.barh(desv.index[::-1], desv.values[::-1], color="#EF4444", height=0.55)
        for b in bars:
            w = b.get_width()
            ax.text(w + (max(desv.values)*0.02), b.get_y() + b.get_height()/2., f"+{int(w):,}", ha='left', va='center', fontsize=8, fontweight='bold', color='#1E293B')
        ax.set_xlabel("Sobrecosto en Horas (Real - Oferta)", fontsize=8.5, color='#475569')

    plt.tight_layout()
    buf = io.BytesIO()
    plt.savefig(buf, format='png', dpi=160)
    plt.close(fig)
    buf.seek(0)
    return buf

def generar_presentacion_powerpoint(metrics, hoja_nombre):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]
    
    # Slide 1: Portada
    slide1 = prs.slides.add_slide(blank_layout)
    bg1 = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    bg1.fill.solid()
    bg1.fill.fore_color.rgb = RGBColor(15, 23, 42)
    bg1.line.fill.background()
    
    tb = slide1.shapes.add_textbox(Inches(1.2), Inches(2.2), Inches(11), Inches(2))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "INFORME EJECUTIVO DE CONTROL OPERACIONAL"
    p.font.size = Pt(32)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 255, 255)
    
    p2 = tf.add_paragraph()
    p2.text = f"Balance de Horas Hombre (HH), Costos y Rendimientos | Proyecto: {hoja_nombre}"
    p2.font.size = Pt(17)
    p2.font.color.rgb = RGBColor(148, 163, 184)
    
    tb_meta = slide1.shapes.add_textbox(Inches(1.2), Inches(5.5), Inches(10), Inches(1))
    tf_m = tb_meta.text_frame
    pm = tf_m.paragraphs[0]
    pm.text = f"Fecha de emisión: {datetime.now().strftime('%d/%m/%Y')}   |   Gerencia de Métodos y Procesos   |   Confidencial"
    pm.font.size = Pt(11)
    pm.font.color.rgb = RGBColor(100, 116, 139)
    
    def agregar_slide_ejecutiva(titulo_slide, subtitulo, chart_buf, kpis_list, notas_default):
        slide = prs.slides.add_slide(blank_layout)
        
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
        bg.fill.solid()
        bg.fill.fore_color.rgb = RGBColor(248, 250, 252)
        bg.line.fill.background()
        
        top_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(1.1))
        top_bar.fill.solid()
        top_bar.fill.fore_color.rgb = RGBColor(15, 23, 42)
        top_bar.line.fill.background()
        
        tb_t = slide.shapes.add_textbox(Inches(0.8), Inches(0.15), Inches(11.5), Inches(0.8))
        tf_t = tb_t.text_frame
        p_t = tf_t.paragraphs[0]
        p_t.text = titulo_slide
        p_t.font.size = Pt(20)
        p_t.font.bold = True
        p_t.font.color.rgb = RGBColor(255, 255, 255)
        
        p_sub = tf_t.add_paragraph()
        p_sub.text = subtitulo
        p_sub.font.size = Pt(11)
        p_sub.font.color.rgb = RGBColor(148, 163, 184)
        
        slide.shapes.add_picture(chart_buf, Inches(0.8), Inches(1.5), width=Inches(7.6))
        
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(8.7), Inches(1.5), Inches(3.9), Inches(2.5))
        card.fill.solid()
        card.fill.fore_color.rgb = RGBColor(255, 255, 255)
        card.line.color.rgb = RGBColor(203, 213, 225)
        
        tb_k = slide.shapes.add_textbox(Inches(8.85), Inches(1.6), Inches(3.6), Inches(2.3))
        tf_k = tb_k.text_frame
        tf_k.word_wrap = True
        p_kh = tf_k.paragraphs[0]
        p_kh.text = "INDICADORES CLAVE"
        p_kh.font.size = Pt(12)
        p_kh.font.bold = True
        p_kh.font.color.rgb = RGBColor(30, 41, 59)
        
        for k_title, k_val, k_pct in kpis_list:
            pk = tf_k.add_paragraph()
            pk.text = f"• {k_title}: {k_val} ({k_pct})"
            pk.font.size = Pt(10.5)
            pk.font.color.rgb = RGBColor(71, 85, 105)
            
        box_n = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(8.7), Inches(4.2), Inches(3.9), Inches(2.7))
        box_n.fill.solid()
        box_n.fill.fore_color.rgb = RGBColor(255, 255, 255)
        box_n.line.color.rgb = RGBColor(203, 213, 225)
        
        tb_n = slide.shapes.add_textbox(Inches(8.85), Inches(4.3), Inches(3.6), Inches(2.5))
        tf_n = tb_n.text_frame
        tf_n.word_wrap = True
        pn_h = tf_n.paragraphs[0]
        pn_h.text = "NOTAS Y OBSERVACIONES"
        pn_h.font.size = Pt(12)
        pn_h.font.bold = True
        pn_h.font.color.rgb = RGBColor(30, 41, 59)
        
        for nota in notas_default:
            pn = tf_n.add_paragraph()
            pn.text = f"- {nota}"
            pn.font.size = Pt(10)
            pn.font.color.rgb = RGBColor(100, 116, 139)

    # Slide 2: Horas
    buf_h = render_chart_image("global_horas", metrics)
    agregar_slide_ejecutiva(
        "BALANCE GLOBAL DE HORAS HOMBRE",
        "Comparativa: Ofertadas vs Estimadas (HdR) vs Reales Ejecutadas",
        buf_h,
        [
            ("Est. vs Oferta", f"+{int(metrics['v_eo']):,} HH", f"{metrics['p_eo']:.1f}%"),
            ("Real vs Est.", f"+{int(metrics['v_re']):,} HH", f"{metrics['p_re']:.1f}%"),
            ("Real vs Oferta", f"+{int(metrics['v_ro']):,} HH", f"{metrics['p_ro']:.1f}%"),
        ],
        [
            f"Ofertadas: {int(metrics['h_ofe']):,} HH | Reales: {int(metrics['h_real']):,} HH",
            "El desvío global refleja mayores requerimientos en fase de taller.",
            "[Haga clic aquí para editar o agregar notas específicas]"
        ]
    )
    
    # Slide 3: Costos
    buf_c = render_chart_image("global_costos", metrics)
    agregar_slide_ejecutiva(
        "IMPACTO FINANCIERO Y COSTOS DE FABRICACIÓN",
        "Cálculo consolidado a partir de tarifas horarias de FABRICOST",
        buf_c,
        [
            ("Est. vs Oferta", f"+{formato_m_k(metrics['vc_eo'])}", f"{metrics['pc_eo']:.1f}%"),
            ("Real vs Est.", f"+{formato_m_k(metrics['vc_re'])}", f"{metrics['pc_re']:.1f}%"),
            ("Real vs Oferta", f"+{formato_m_k(metrics['vc_ro'])}", f"{metrics['pc_ro']:.1f}%"),
        ],
        [
            f"Costo Oferta: {formato_m_k(metrics['c_ofe'])} | Real: {formato_m_k(metrics['c_real'])}",
            "Tarifas operativas y CIF aplicadas por disciplina.",
            "[Haga clic aquí para editar o agregar notas específicas]"
        ]
    )

    # Slide 4: Rendimientos
    buf_r = render_chart_image("rendimientos", metrics)
    agregar_slide_ejecutiva(
        "RENDIMIENTO Y PRODUCTIVIDAD (Kg / HH)",
        "Evolución del ratio de kilos procesados por hora hombre",
        buf_r,
        [
            ("Rend. Oferta", f"{metrics['rend_ofe']:.2f} Kg/HH", "Base"),
            ("Rend. Estimado", f"{metrics['rend_est']:.2f} Kg/HH", f"{metrics['p_rend_eo']:.1f}%"),
            ("Rend. Real", f"{metrics['rend_real']:.2f} Kg/HH", f"{metrics['p_rend_ro']:.1f}%"),
        ],
        [
            f"Variación Real vs Oferta: {metrics['d_rend_ro']:.2f} Kg/HH",
            "Una reducción en el ratio indica mayor consumo de horas por tonelada.",
            "[Haga clic aquí para editar o agregar notas específicas]"
        ]
    )

    # Slide 5: Puestos Críticos
    if 'desvios_tarea' in metrics and not metrics['desvios_tarea'].empty:
        buf_d = render_chart_image("top_desvios", metrics)
        agregar_slide_ejecutiva(
            "PUESTOS CRÍTICOS CON MAYOR SOBRECOSTO",
            "Top de disciplinas con mayor desviación de horas (Real - Oferta)",
            buf_d,
            [
                (str(metrics['desvios_tarea'].index[0]), f"+{int(metrics['desvios_tarea'].iloc[0]):,} HH", "Mayor desvío"),
                (str(metrics['desvios_tarea'].index[1]), f"+{int(metrics['desvios_tarea'].iloc[1]):,} HH", "Puesto 2"),
                (str(metrics['desvios_tarea'].index[2]), f"+{int(metrics['desvios_tarea'].iloc[2]):,} HH", "Puesto 3"),
            ],
            [
                "Concentración principal del desvío en talleres mecánicos y estructura.",
                "Puntos focales para análisis de métodos en siguientes proyectos.",
                "[Haga clic aquí para editar o agregar notas específicas]"
            ]
        )
        
    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out

def generar_informe_pdf(metrics, hoja_nombre):
    pdf_buf = io.BytesIO()
    doc = SimpleDocTemplate(pdf_buf, pagesize=landscape(letter), rightMargin=25, leftMargin=25, topMargin=25, bottomMargin=25)
    styles = getSampleStyleSheet()
    
    t_style = ParagraphStyle(name='TitleStyle', parent=styles['Heading1'], fontSize=16, textColor=colors.HexColor('#0F172A'), spaceAfter=3)
    sub_style = ParagraphStyle(name='SubStyle', parent=styles['Normal'], fontSize=9, textColor=colors.HexColor('#64748B'), spaceAfter=12)
    h2_style = ParagraphStyle(name='H2Style', parent=styles['Heading2'], fontSize=11, textColor=colors.HexColor('#1E3A8A'), spaceBefore=6, spaceAfter=4)
    
    story = []
    story.append(Paragraph(f"INFORME GERENCIAL DE CONTROL OPERACIONAL - {hoja_nombre.upper()}", t_style))
    story.append(Paragraph(f"Fecha: {datetime.now().strftime('%d/%m/%Y')} | Departamento de Métodos y Procesos | Reporte Consolidado", sub_style))
    
    story.append(Paragraph("1. Resumen Consolidado de Desvíos (Oferta vs Estimado vs Real)", h2_style))
    tabla_kpi_data = [
        ["Dimensión", "Ofertado", "Estimado (HdR)", "Real Ejecutado", "Desvío Real vs Oferta", "% Desvío"],
        ["Horas Hombre (HH)", f"{int(metrics['h_ofe']):,}", f"{int(metrics['h_est']):,}", f"{int(metrics['h_real']):,}", f"+{int(metrics['v_ro']):,}", f"{metrics['p_ro']:.2f}%"],
        ["Costo Total ($)", formato_m_k(metrics['c_ofe']), formato_m_k(metrics['c_est']), formato_m_k(metrics['c_real']), f"+{formato_m_k(metrics['vc_ro'])}", f"{metrics['pc_ro']:.2f}%"],
        ["Rendimiento (Kg/HH)", f"{metrics['rend_ofe']:.2f}", f"{metrics['rend_est']:.2f}", f"{metrics['rend_real']:.2f}", f"{metrics['d_rend_ro']:.2f}", f"{metrics['p_rend_ro']:.2f}%"]
    ]
    t_kpi = Table(tabla_kpi_data, colWidths=[140, 95, 105, 105, 115, 95])
    t_kpi.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#0F172A')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 8.5),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CBD5E1')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F8FAFC')]),
        ('ALIGN', (1, 0), (-1, -1), 'CENTER'),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ]))
    story.append(t_kpi)
    story.append(Spacer(1, 14))
    
    story.append(Paragraph("2. Comparativa Gráfica de Métricas Clave", h2_style))
    buf_gh = render_chart_image("global_horas", metrics)
    buf_gc = render_chart_image("global_costos", metrics)
    
    img_h = RLImage(buf_gh, width=4.8*72, height=2.6*72)
    img_c = RLImage(buf_gc, width=4.8*72, height=2.6*72)
    
    t_imgs = Table([[img_h, img_c]], colWidths=[330, 330])
    t_imgs.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ]))
    story.append(t_imgs)
    
    doc.build(story)
    pdf_buf.seek(0)
    return pdf_buf

if uploaded_file:
    file_bytes = io.BytesIO(uploaded_file.read())
    wb_check = openpyxl.load_workbook(file_bytes, read_only=True)
    hojas = wb_check.sheetnames
    wb_check.close()
    
    file_bytes.seek(0)
    df_tarifas = leer_fabricost(file_bytes)
    
    hojas_datos = [h for h in hojas if h != "FABRICOST"]
    hoja = st.sidebar.selectbox("Hoja de Proyecto activa:", hojas_datos, index=0)
    
    file_bytes.seek(0)
    df_raw = leer_matriz_exacta(file_bytes, hoja)
    df_long, df_pesos = procesar_datos(df_raw, df_tarifas)
    
    # --- FILTROS ---
    st.sidebar.markdown("---")
    st.sidebar.subheader("🔍 Filtros Dinámicos")
    
    lista_recursos = sorted(df_long['Recurso'].unique().tolist())
    filtro_recurso = st.sidebar.multiselect("RECURSO:", options=lista_recursos, default=[])
    
    lista_gdf = sorted(df_long['GDF'].dropna().unique().tolist())
    filtro_gdf = st.sidebar.multiselect("GDF:", options=lista_gdf, default=[])
    
    df_f = df_long.copy()
    df_p_f = df_pesos.copy()
    if filtro_recurso:
        df_f = df_f[df_f['Recurso'].isin(filtro_recurso)]
    if filtro_gdf:
        df_f = df_f[df_f['GDF'].isin(filtro_gdf)]
        df_p_f = df_p_f[df_p_f['GDF'].isin(filtro_gdf)]
        
    tot_horas = df_f.groupby('Estado')['Horas'].sum().to_dict()
    h_ofe = tot_horas.get("OFERTADAS", 0)
    h_est = tot_horas.get("ESTIMADAS", 0)
    h_real = tot_horas.get("REALES", 0)
    
    tot_costos = df_f.groupby('Estado')['Costo'].sum().to_dict()
    c_ofe = tot_costos.get("OFERTADAS", 0)
    c_est = tot_costos.get("ESTIMADAS", 0)
    c_real = tot_costos.get("REALES", 0)
    
    tot_peso = df_p_f.groupby('Estado')['Peso'].sum().to_dict()
    p_ofe = tot_peso.get("OFERTADAS", 0)
    p_est = tot_peso.get("ESTIMADAS", 0)
    p_real = tot_peso.get("REALES", 0)
    
    rend_ofe = (p_ofe / h_ofe) if h_ofe > 0 else 0
    rend_est = (p_est / h_est) if h_est > 0 else 0
    rend_real = (p_real / h_real) if h_real > 0 else 0
    
    v_eo = h_est - h_ofe
    p_eo = (v_eo / h_ofe * 100) if h_ofe else 0
    v_re = h_real - h_est
    p_re = (v_re / h_est * 100) if h_est else 0
    v_ro = h_real - h_ofe
    p_ro = (v_ro / h_ofe * 100) if h_ofe else 0
    
    vc_eo = c_est - c_ofe
    pc_eo = (vc_eo / c_ofe * 100) if c_ofe else 0
    vc_re = c_real - c_est
    pc_re = (vc_re / c_est * 100) if c_est else 0
    vc_ro = c_real - c_ofe
    pc_ro = (vc_ro / c_ofe * 100) if c_ofe else 0
    
    d_rend_eo = rend_est - rend_ofe
    p_rend_eo = (d_rend_eo / rend_ofe * 100) if rend_ofe else 0
    d_rend_re = rend_real - rend_est
    p_rend_re = (d_rend_re / rend_est * 100) if rend_est else 0
    d_rend_ro = rend_real - rend_ofe
    p_rend_ro = (d_rend_ro / rend_ofe * 100) if rend_ofe else 0

    piv_desv = df_f.pivot_table(index='Recurso', columns='Estado', values='Horas', aggfunc='sum').fillna(0)
    desvios_tarea = pd.Series()
    if 'OFERTADAS' in piv_desv.columns and 'REALES' in piv_desv.columns:
        desvios_tarea = (piv_desv['REALES'] - piv_desv['OFERTADAS']).sort_values(ascending=False)

    dict_metrics = {
        'h_ofe': h_ofe, 'h_est': h_est, 'h_real': h_real,
        'v_eo': v_eo, 'p_eo': p_eo, 'v_re': v_re, 'p_re': p_re, 'v_ro': v_ro, 'p_ro': p_ro,
        'c_ofe': c_ofe, 'c_est': c_est, 'c_real': c_real,
        'vc_eo': vc_eo, 'pc_eo': pc_eo, 'vc_re': vc_re, 'pc_re': pc_re, 'vc_ro': vc_ro, 'pc_ro': pc_ro,
        'rend_ofe': rend_ofe, 'rend_est': rend_est, 'rend_real': rend_real,
        'd_rend_eo': d_rend_eo, 'p_rend_eo': p_rend_eo, 'd_rend_re': d_rend_re, 'p_rend_re': p_rend_re,
        'd_rend_ro': d_rend_ro, 'p_rend_ro': p_rend_ro,
        'desvios_tarea': desvios_tarea
    }

    # --- BOTONES EN LA BARRA LATERAL ---
    st.sidebar.markdown("---")
    st.sidebar.subheader("📥 Exportación Ejecutiva")
    
    pptx_bytes = generar_presentacion_powerpoint(dict_metrics, hoja)
    pdf_bytes = generar_informe_pdf(dict_metrics, hoja)
    
    st.sidebar.download_button(
        label="📊 Descargar Presentación (.pptx)",
        data=pptx_bytes,
        file_name=f"Reporte_Ejecutivo_{hoja}.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )
    st.sidebar.download_button(
        label="📄 Descargar Informe (.pdf)",
        data=pdf_bytes,
        file_name=f"Informe_Ejecutivo_{hoja}.pdf",
        mime="application/pdf"
    )

    color_map = {"OFERTADAS": "#F59E0B", "ESTIMADAS": "#06B6D4", "REALES": "#6366F1"}
    orden_estados = ["OFERTADAS", "ESTIMADAS", "REALES"]

    def crear_cascada(df_base, col_agrupacion, metrica, estado_fin, estado_ini, titulo):
        piv = df_base.pivot_table(index=col_agrupacion, columns='Estado', values=metrica, aggfunc='sum').fillna(0)
        if estado_fin in piv.columns and estado_ini in piv.columns:
            piv['Desvio'] = piv[estado_fin] - piv[estado_ini]
            piv = piv.sort_values(by='Desvio', ascending=False).reset_index()
            
            y_vals = piv['Desvio'].tolist()
            y_total = sum(y_vals)
            text_labels = [formato_m_k(v) if metrica == 'Costo' else f"{int(v):,}" for v in y_vals] + [formato_m_k(y_total) if metrica == 'Costo' else f"{int(y_total):,}"]
            
            fig = go.Figure(go.Waterfall(
                orientation="v",
                measure=["relative"] * len(piv) + ["total"],
                x=piv[col_agrupacion].tolist() + ["Total"],
                y=y_vals + [y_total],
                text=text_labels,
                textposition="outside",
                decreasing={"marker": {"color": "#10B981"}},
                increasing={"marker": {"color": "#EF4444"}},
                totals={"marker": {"color": "#06B6D4"}}
            ))
            fig.update_layout(title=titulo, height=420, xaxis_tickangle=-45, yaxis_title="")
            return fig
        return None

    # --- PESTAÑAS ---
    tabs = st.tabs([
        "HORAS 1", "HORAS 2", "ESTIMADAS VS OFERTADAS", "REALES VS ESTIMADAS", "REALES VS OFERTADAS",
        "COSTOS", "COSTOS 2", "COSTOS ESTIM VS OFERT", "COSTOS REAL VS ESTIM", "COSTOS REAL VS OF",
        "RENDIMIENTOS"
    ])
    
    with tabs[0]:
        st.subheader("HORAS")
        df_bh = pd.DataFrame({"Estado": orden_estados, "Horas": [h_ofe, h_est, h_real]})
        fig_h = px.bar(df_bh, x="Estado", y="Horas", color="Estado", text_auto=",.0f", color_discrete_map=color_map)
        fig_h.update_layout(height=400, showlegend=False, xaxis_title="", yaxis_title="Horas")
        st.plotly_chart(fig_h, use_container_width=True)
        
        c1, c2, c3 = st.columns(3)
        c1.metric("VARIACION HH ESTIMADO VS OFERTADO", f"{int(v_eo):,}", f"{p_eo:.2f}%")
        c2.metric("VARIACION HH REAL VS ESTIMADO", f"{int(v_re):,}", f"{p_re:.2f}%")
        c3.metric("VARIACION HH REALES VS OFERTADAS", f"{int(v_ro):,}", f"{p_ro:.2f}%")

    with tabs[1]:
        st.subheader("HORAS POR TAREA")
        df_t = df_f.groupby(['Recurso', 'Estado'])['Horas'].sum().reset_index()
        fig_t = px.bar(df_t, x="Recurso", y="Horas", color="Estado", barmode="group", category_orders={"Estado": orden_estados}, color_discrete_map=color_map)
        fig_t.update_layout(height=420, xaxis_tickangle=-45, yaxis_title="", legend_title="")
        st.plotly_chart(fig_t, use_container_width=True)
        
        st.subheader("HORAS POR GDF")
        df_g = df_f.groupby(['GDF', 'Estado'])['Horas'].sum().reset_index()
        fig_g = px.bar(df_g, x="GDF", y="Horas", color="Estado", barmode="group", category_orders={"Estado": orden_estados}, color_discrete_map=color_map)
        fig_g.update_layout(height=420, yaxis_title="", legend_title="")
        st.plotly_chart(fig_g, use_container_width=True)

    with tabs[2]:
        fig1 = crear_cascada(df_f, 'GDF', 'Horas', 'ESTIMADAS', 'OFERTADAS', "HH ESTIMADO VS OFERTADO POR GDF")
        if fig1: st.plotly_chart(fig1, use_container_width=True)
        fig2 = crear_cascada(df_f, 'Recurso', 'Horas', 'ESTIMADAS', 'OFERTADAS', "HH ESTIMADO VS OFERTADO POR TAREA")
        if fig2: st.plotly_chart(fig2, use_container_width=True)

    with tabs[3]:
        fig1 = crear_cascada(df_f, 'GDF', 'Horas', 'REALES', 'ESTIMADAS', "HH REALES VS ESTIMADAS POR GDF")
        if fig1: st.plotly_chart(fig1, use_container_width=True)
        fig2 = crear_cascada(df_f, 'Recurso', 'Horas', 'REALES', 'ESTIMADAS', "HH REALES VS ESTIMADAS POR TAREA")
        if fig2: st.plotly_chart(fig2, use_container_width=True)

    with tabs[4]:
        fig1 = crear_cascada(df_f, 'GDF', 'Horas', 'REALES', 'OFERTADAS', "HH REALES VS OFERTADAS POR GDF")
        if fig1: st.plotly_chart(fig1, use_container_width=True)
        fig2 = crear_cascada(df_f, 'Recurso', 'Horas', 'REALES', 'OFERTADAS', "VARIACION HH REALES VS OFERTADAS by RECURSO")
        if fig2: st.plotly_chart(fig2, use_container_width=True)

    with tabs[5]:
        st.subheader("COSTOS")
        df_bc = pd.DataFrame({"Estado": orden_estados, "Costo": [c_ofe, c_est, c_real]})
        df_bc["Etiqueta"] = df_bc["Costo"].apply(formato_m_k)
        fig_c = px.bar(df_bc, x="Estado", y="Costo", color="Estado", text="Etiqueta", color_discrete_map=color_map)
        fig_c.update_layout(height=400, showlegend=False, xaxis_title="", yaxis_title="Costo Total")
        st.plotly_chart(fig_c, use_container_width=True)
        
        c1, c2, c3 = st.columns(3)
        c1.metric("COSTOS ESTIMADOS VS OFERTADOS", formato_m_k(vc_eo), f"{pc_eo:.2f}%")
        c2.metric("COSTOS REALES VS ESTIMADOS", formato_m_k(vc_re), f"{pc_re:.2f}%")
        c3.metric("COSTOS REALES VS OFERTADOS", formato_m_k(vc_ro), f"{pc_ro:.2f}%")

    with tabs[6]:
        st.subheader("COSTOS POR RECURSO")
        df_cr = df_f.groupby(['Recurso', 'Estado'])['Costo'].sum().reset_index()
        fig_cr = px.bar(df_cr, x="Recurso", y="Costo", color="Estado", barmode="group", category_orders={"Estado": orden_estados}, color_discrete_map=color_map)
        fig_cr.update_layout(height=420, xaxis_tickangle=-45, yaxis_title="", legend_title="")
        st.plotly_chart(fig_cr, use_container_width=True)
        
        st.subheader("COSTOS POR GDF")
        df_cg = df_f.groupby(['GDF', 'Estado'])['Costo'].sum().reset_index()
        fig_cg = px.bar(df_cg, x="GDF", y="Costo", color="Estado", barmode="group", category_orders={"Estado": orden_estados}, color_discrete_map=color_map)
        fig_cg.update_layout(height=420, yaxis_title="", legend_title="")
        st.plotly_chart(fig_cg, use_container_width=True)

    with tabs[7]:
        fig1 = crear_cascada(df_f, 'GDF', 'Costo', 'ESTIMADAS', 'OFERTADAS', "COSTOS ESTIMADOS VS OFERTADOS POR GDF")
        if fig1: st.plotly_chart(fig1, use_container_width=True)
        fig2 = crear_cascada(df_f, 'Recurso', 'Costo', 'ESTIMADAS', 'OFERTADAS', "COSTOS ESTIMADOS VS OFERTADOS POR TAREA")
        if fig2: st.plotly_chart(fig2, use_container_width=True)

    with tabs[8]:
        fig1 = crear_cascada(df_f, 'GDF', 'Costo', 'REALES', 'ESTIMADAS', "COSTOS REALES VS ESTIMADOS POR GDF")
        if fig1: st.plotly_chart(fig1, use_container_width=True)
        fig2 = crear_cascada(df_f, 'Recurso', 'Costo', 'REALES', 'ESTIMADAS', "COSTOS REALES VS ESTIMADOS POR RECURSO")
        if fig2: st.plotly_chart(fig2, use_container_width=True)

    with tabs[9]:
        fig1 = crear_cascada(df_f, 'GDF', 'Costo', 'REALES', 'OFERTADAS', "COSTOS REALES VS OFERTADOS POR GDF")
        if fig1: st.plotly_chart(fig1, use_container_width=True)
        fig2 = crear_cascada(df_f, 'Recurso', 'Costo', 'REALES', 'OFERTADAS', "COSTOS REALES VS OFERTADOS POR RECURSO")
        if fig2: st.plotly_chart(fig2, use_container_width=True)

    with tabs[10]:
        st.subheader("RENDIMIENTOS (Peso / HH)")
        df_rend = pd.DataFrame({
            "Estado": ["OFERTADO", "ESTIMADO", "REAL"],
            "Rendimiento": [rend_ofe, rend_est, rend_real]
        })
        fig_r = px.bar(df_rend, x="Estado", y="Rendimiento", color="Estado", text=df_rend["Rendimiento"].apply(lambda v: f"{v:.2f}"), color_discrete_map={"OFERTADO": "#F59E0B", "ESTIMADO": "#06B6D4", "REAL": "#6366F1"})
        fig_r.update_layout(height=400, showlegend=False, xaxis_title="", yaxis_title="Kg / HH")
        st.plotly_chart(fig_r, use_container_width=True)
        
        cr1, cr2, cr3 = st.columns(3)
        cr1.metric("REND ESTIMADO VS OFERTADO", f"{d_rend_eo:.2f}", f"{p_rend_eo:.2f}% (%V E-O)")
        cr2.metric("REND REAL VS ESTIM", f"{d_rend_re:.2f}", f"{p_rend_re:.2f}% (%V R-E)")
        cr3.metric("REND REAL VS OFERTADO", f"{d_rend_ro:.2f}", f"{p_rend_ro:.2f}% (%V R-O)")
