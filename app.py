import io
import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
import openai

st.set_page_config(page_title="Generador de Reportes de Métodos y Costos", layout="wide")

st.title("📊 Generador de Reportes Industriales y Métodos")
st.write("Carga de presupuestos, matrices de horas hombre (HH) y balances de proyecto.")

# Configuración en barra lateral
api_key = st.sidebar.text_input("OpenAI API Key", type="password", help="Tu API Key no se almacena.")
st.sidebar.markdown("---")

uploaded_file = st.file_uploader("Arrastra tu archivo Excel aquí", type=["xlsx", "xls"])

def limpiar_matriz_industrial(df_raw):
    """
    Detecta automáticamente la fila de encabezados reales y resuelve celdas combinadas.
    """
    header_idx = None
    for idx in range(min(10, len(df_raw))):
        row_vals = [str(v) for v in df_raw.iloc[idx].values if pd.notna(v)]
        row_str = " ".join(row_vals)
        if any(k in row_str for k in ["Peso", "HH", "RECURSO", "GDF", "Oferta", "Subconjunto"]):
            header_idx = idx
            break
            
    if header_idx is not None:
        df = df_raw.iloc[header_idx + 1:].copy()
        df.columns = df_raw.iloc[header_idx].values
    else:
        df = df_raw.copy()

    df.columns = [str(c).strip().replace("\n", " ") if (pd.notna(c) and str(c).strip() != "") else f"Col_{i}" for i, c in enumerate(df.columns)]
    
    df = df.dropna(how='all', axis=1)
    df = df.dropna(how='all', axis=0)

    for col in df.columns[:3]:
        df[col] = df[col].ffill()
        
    return df

def analizar_con_ia(contexto_proyecto, key):
    client = openai.OpenAI(api_key=key)
    prompt = f"""
    Eres un ingeniero experto en Métodos, Procesos Industriales y Costos Operativos.
    Analiza la matriz de horas hombre (HH) y pesos del proyecto:
    
    {contexto_proyecto}
    
    Genera un informe gerencial estructurado estrictamente en:
    1. Resumen Ejecutivo (Visión general del estado del proyecto: desvíos generales entre Oferta, Estimado y Real).
    2. Análisis de Desvíos de Recursos Críticos (Compara las principales horas hombre: Estructura, Soldadura, Mecanizado/Tornos, QA-QC, etc.).
    3. Análisis de Pesos (Oferta vs Fabricación).
    4. Conclusiones y Alertas Operativas (Puntos de atención para el responsable de Métodos / Operaciones).
    """
    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[{"role": "user", "content": prompt}],
        temperature=0.2,
        max_tokens=800
    )
    return response.choices[0].message.content

def crear_powerpoint(dict_dfs, analisis_ia):
    prs = Presentation()
    
    # Portada
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = "Informe Ejecutivo de Desvíos y Métodos"
    slide.placeholders[1].text = "Control de Horas Hombre (HH) y Pesos de Proyecto\nConfidencial"
    
    # Diagnóstico IA
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = "Diagnóstico Estratégico y Desvíos"
    tx_box = slide.shapes.add_textbox(Inches(0.6), Inches(1.3), Inches(8.8), Inches(5.5))
    tf = tx_box.text_frame
    tf.word_wrap = True
    tf.text = analisis_ia
    
    # Diapositivas por hoja
    for sheet_name, df_sheet in list(dict_dfs.items())[:3]:
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        slide.shapes.title.text = f"Matriz: {sheet_name}"
        
        df_display = df_sheet.head(7)
        if not df_display.empty:
            cols_to_show = min(len(df_display.columns), 7)
            rows = len(df_display) + 1
            table_shape = slide.shapes.add_table(rows, cols_to_show, Inches(0.5), Inches(1.5), Inches(9.0), Inches(4.5))
            table = table_shape.table
            
            for i, col_name in enumerate(df_display.columns[:cols_to_show]):
                cell = table.cell(0, i)
                cell.text = str(col_name)[:15]
                
            for row_idx, (_, row) in enumerate(df_display.iterrows()):
                for col_idx in range(cols_to_show):
                    cell = table.cell(row_idx + 1, col_idx)
                    val = row.iloc[col_idx]
                    cell.text = str(round(val, 1)) if isinstance(val, (int, float)) and pd.notna(val) else (str(val) if pd.notna(val) else "-")
                    
    pptx_io = io.BytesIO()
    prs.save(pptx_io)
    pptx_io.seek(0)
    return pptx_io

def crear_pdf(dict_dfs, analisis_ia):
    pdf_io = io.BytesIO()
    doc = SimpleDocTemplate(pdf_io, pagesize=landscape(letter), rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle(name='TitleStyle', parent=styles['Heading1'], fontSize=15, textColor=colors.HexColor('#0F172A'))
    h2_style = ParagraphStyle(name='H2Style', parent=styles['Heading2'], fontSize=11, textColor=colors.HexColor('#1E40AF'))
    body_style = ParagraphStyle(name='BodyStyle', parent=styles['Normal'], fontSize=8.5, leading=12)
    
    story = []
    story.append(Paragraph("Informe Gerencial de Métodos, HH y Costos de Fabricación", title_style))
    story.append(Spacer(1, 8))
    
    story.append(Paragraph("1. Análisis Técnico y Desvíos", h2_style))
    story.append(Spacer(1, 4))
    for parrafo in analisis_ia.split("\n"):
        if parrafo.strip():
            story.append(Paragraph(parrafo, body_style))
            story.append(Spacer(1, 2))
    story.append(Spacer(1, 10))
    
    for sheet_name, df_sheet in list(dict_dfs.items())[:2]:
        story.append(Paragraph(f"Resumen de Matriz: {sheet_name}", h2_style))
        story.append(Spacer(1, 4))
        
        cols_to_use = [c for c in df_sheet.columns if not c.startswith("Col_")][:8]
        if not cols_to_use:
            cols_to_use = df_sheet.columns[:8].tolist()
            
        df_sub = df_sheet[cols_to_use].head(8)
        tabla_data = [cols_to_use]
        for _, row in df_sub.iterrows():
            fila = []
            for val in row:
                fila.append(str(round(val, 1)) if isinstance(val, (int, float)) and pd.notna(val) else (str(val)[:15] if pd.notna(val) else "-"))
            tabla_data.append(fila)
            
        t = Table(tabla_data)
        t.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#E2E8F0')),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 7.5),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#94A3B8')),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
            ('TOPPADDING', (0, 0), (-1, -1), 3),
        ]))
        story.append(t)
        story.append(Spacer(1, 8))
        
    doc.build(story)
    pdf_io.seek(0)
    return pdf_io

if uploaded_file is not None:
    try:
        excel_file = pd.ExcelFile(uploaded_file)
        nombres_hojas = excel_file.sheet_names
        
        st.success(f"Archivo cargado. Se detectaron {len(nombres_hojas)} hojas en el libro.")
        
        hojas_seleccionadas = st.multiselect(
            "Selecciona qué hojas quieres incluir en el análisis:",
            options=nombres_hojas,
            default=[nombres_hojas[1]] if len(nombres_hojas) > 1 else [nombres_hojas[0]]
        )
        
        if not hojas_seleccionadas:
            st.warning("Selecciona al menos una hoja para continuar.")
        else:
            dict_dfs = {}
            tabs = st.tabs(hojas_seleccionadas)
            for i, sheet in enumerate(hojas_seleccionadas):
                df_raw = pd.read_excel(uploaded_file, sheet_name=sheet, header=None)
                df_limpio = limpiar_matriz_industrial(df_raw)
                dict_dfs[sheet] = df_limpio
                
                with tabs[i]:
                    st.write(f"**Vista previa limpia de {sheet}** ({len(df_limpio)} filas detectadas)")
                    st.dataframe(df_limpio.head(12))
            
            if not api_key:
                st.warning("Ingresa tu OpenAI API Key en la barra lateral para habilitar la generación.")
            else:
                if st.button("🚀 Generar Informe de Proyecto (.pptx y .pdf)"):
                    with st.spinner("Analizando desvíos de HH, pesos y generando entregables..."):
                        contexto = ""
                        for s_name, s_df in dict_dfs.items():
                            contexto += f"\n=== MATRIZ: {s_name} ===\n"
                            filas_totales = s_df[s_df.apply(lambda r: r.astype(str).str.contains("Total|OFERTADO|REAL|ESTIMADO", case=False).any(), axis=1)]
                            if not filas_totales.empty:
                                contexto += f"Resumen de Totales y Bloques:\n{filas_totales.to_string()}\n"
                            else:
                                contexto += f"Datos:\n{s_df.head(10).to_string()}\n"
                        
                        analisis = analizar_con_ia(contexto, api_key)
                        pptx_file = crear_powerpoint(dict_dfs, analisis)
                        pdf_file = crear_pdf(dict_dfs, analisis)
                        
                        st.success("¡Informe y presentación generados exitosamente!")
                        col1, col2 = st.columns(2)
                        with col1:
                            st.download_button(
                                label="📥 Descargar Presentación (.pptx)",
                                data=pptx_file,
                                file_name="Informe_Metodos_HH.pptx",
                                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                            )
                        with col2:
                            st.download_button(
                                label="📥 Descargar Reporte en PDF (Horizontal)",
                                data=pdf_file,
                                file_name="Informe_Metodos_HH.pdf",
                                mime="application/pdf"
                            )
    except Exception as e:
        st.error(f"Error al procesar las matrices del Excel: {e}")
