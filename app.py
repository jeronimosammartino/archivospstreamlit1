import io
import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
import openai

st.set_page_config(page_title="Generador de Reportes Corporativos", layout="wide")

st.title("📊 Generador Automatizado de Reportes")
st.write("Sube tu archivo Excel (.xlsx) con múltiples hojas para procesar los datos y generar presentaciones y reportes ejecutivos.")

# Configuración en barra lateral
api_key = st.sidebar.text_input("OpenAI API Key", type="password", help="Tu API Key no se guarda en el servidor.")
st.sidebar.markdown("---")

uploaded_file = st.file_uploader("Arrastra tu archivo Excel aquí", type=["xlsx", "xls"])

def analizar_con_ia(contexto_hojas, key):
    client = openai.OpenAI(api_key=key)
    prompt = f"""
    Eres un analista senior de costos, operaciones y proyectos de ingeniería.
    Analiza la información de las siguientes hojas del proyecto:
    
    {contexto_hojas}
    
    Estructura tu respuesta estrictamente en:
    1. Resumen Ejecutivo (Visión general del costeo/proyecto en 3-4 líneas).
    2. Conclusiones Clave de Costos y Recursos (3 a 5 puntos con viñetas concisas).
    3. Alertas o Recomendaciones Operativas (2 a 3 puntos de acción).
    """
    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[{"role": "user", "content": prompt}],
        temperature=0.2,
        max_tokens=600
    )
    return response.choices[0].message.content

def crear_powerpoint(dict_dfs, analisis_ia):
    prs = Presentation()
    
    # Slide 1: Portada
    slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(slide_layout)
    slide.shapes.title.text = "Informe Ejecutivo de Costos y Proyecto"
    slide.placeholders[1].text = "Análisis consolidado de estructura de costos\nConfidencial"
    
    # Slide 2: Conclusiones Ejecutivas
    slide_layout = prs.slide_layouts[5]
    slide = prs.slides.add_slide(slide_layout)
    slide.shapes.title.text = "Diagnóstico Estratégico y Conclusiones"
    tx_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.5), Inches(8.4), Inches(5))
    tf = tx_box.text_frame
    tf.word_wrap = True
    tf.text = analisis_ia
    
    # Slides por cada hoja seleccionada (máximo 4 hojas)
    for sheet_name, df_sheet in list(dict_dfs.items())[:4]:
        slide = prs.slides.add_slide(slide_layout)
        slide.shapes.title.text = f"Detalle: {sheet_name}"
        
        df_display = df_sheet.dropna(how='all').head(6)
        if not df_display.empty:
            rows, cols = len(df_display) + 1, min(len(df_display.columns), 6)
            table_shape = slide.shapes.add_table(rows, cols, Inches(0.8), Inches(1.8), Inches(8.4), Inches(3.8))
            table = table_shape.table
            
            # Encabezados
            for i, col_name in enumerate(df_display.columns[:cols]):
                cell = table.cell(0, i)
                cell.text = str(col_name)
                
            # Datos
            for row_idx, (_, row) in enumerate(df_display.iterrows()):
                for col_idx in range(cols):
                    cell = table.cell(row_idx + 1, col_idx)
                    val = row.iloc[col_idx]
                    cell.text = str(round(val, 2)) if isinstance(val, float) else str(val)
                    
    pptx_io = io.BytesIO()
    prs.save(pptx_io)
    pptx_io.seek(0)
    return pptx_io

def crear_pdf(dict_dfs, analisis_ia):
    pdf_io = io.BytesIO()
    doc = SimpleDocTemplate(pdf_io, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle(name='TitleStyle', parent=styles['Heading1'], fontSize=16, textColor=colors.HexColor('#0F172A'))
    h2_style = ParagraphStyle(name='H2Style', parent=styles['Heading2'], fontSize=12, textColor=colors.HexColor('#1E40AF'))
    body_style = ParagraphStyle(name='BodyStyle', parent=styles['Normal'], fontSize=9, leading=13)
    
    story = []
    story.append(Paragraph("Informe Ejecutivo de Costos Operativos", title_style))
    story.append(Spacer(1, 10))
    
    story.append(Paragraph("1. Conclusiones y Diagnóstico Operativo", h2_style))
    story.append(Spacer(1, 6))
    for parrafo in analisis_ia.split("\n"):
        if parrafo.strip():
            story.append(Paragraph(parrafo, body_style))
            story.append(Spacer(1, 3))
    story.append(Spacer(1, 12))
    
    for sheet_name, df_sheet in list(dict_dfs.items())[:3]:
        df_clean = df_sheet.dropna(how='all').head(8)
        if not df_clean.empty:
            story.append(Paragraph(f"Detalle de Hoja: {sheet_name}", h2_style))
            story.append(Spacer(1, 4))
            
            cols_to_use = df_clean.columns[:5].tolist()
            tabla_data = [cols_to_use]
            for _, row in df_clean[cols_to_use].iterrows():
                fila = []
                for val in row:
                    fila.append(str(round(val, 2)) if isinstance(val, float) else str(val))
                tabla_data.append(fila)
                
            t = Table(tabla_data)
            t.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#F1F5F9')),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, -1), 8),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CBD5E1')),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
                ('TOPPADDING', (0, 0), (-1, -1), 4),
            ]))
            story.append(t)
            story.append(Spacer(1, 10))
            
    doc.build(story)
    pdf_io.seek(0)
    return pdf_io

if uploaded_file is not None:
    try:
        excel_file = pd.ExcelFile(uploaded_file)
        nombres_hojas = excel_file.sheet_names
        
        st.success(f"Archivo cargado. Se detectaron {len(nombres_hojas)} hojas en el libro.")
        
        hojas_seleccionadas = st.multiselect(
            "Selecciona qué hojas quieres incluir en el reporte:",
            options=nombres_hojas,
            default=nombres_hojas[:3] if len(nombres_hojas) >= 3 else nombres_hojas
        )
        
        if not hojas_seleccionadas:
            st.warning("Selecciona al menos una hoja para continuar.")
        else:
            dict_dfs = {}
            tabs = st.tabs(hojas_seleccionadas)
            for i, sheet in enumerate(hojas_seleccionadas):
                df_temp = pd.read_excel(uploaded_file, sheet_name=sheet)
                dict_dfs[sheet] = df_temp
                with tabs[i]:
                    st.write(f"**Vista previa de {sheet}** ({len(df_temp)} filas, {len(df_temp.columns)} columnas)")
                    st.dataframe(df_temp.head(5))
            
            if not api_key:
                st.warning("Ingresa tu OpenAI API Key en la barra lateral para habilitar la generación.")
            else:
                if st.button("🚀 Generar Reportes Multi-Hoja (.pptx y .pdf)"):
                    with st.spinner("Consolidando hojas y generando análisis con IA..."):
                        contexto = ""
                        for s_name, s_df in dict_dfs.items():
                            contexto += f"\n--- HOJA: {s_name} ---\n"
                            contexto += f"Columnas: {list(s_df.columns)}\n"
                            contexto += f"Muestra de datos:\n{s_df.head(6).to_string()}\n"
                            num_cols = s_df.select_dtypes(include='number')
                            if not num_cols.empty:
                                contexto += f"Totales / Promedios:\n{num_cols.agg(['sum', 'mean']).to_string()}\n"
                        
                        analisis = analizar_con_ia(contexto, api_key)
                        pptx_file = crear_powerpoint(dict_dfs, analisis)
                        pdf_file = crear_pdf(dict_dfs, analisis)
                        
                        st.success("¡Reportes generados exitosamente!")
                        col1, col2 = st.columns(2)
                        with col1:
                            st.download_button(
                                label="📥 Descargar PowerPoint (.pptx)",
                                data=pptx_file,
                                file_name="Reporte_MultiHoja.pptx",
                                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                            )
                        with col2:
                            st.download_button(
                                label="📥 Descargar Reporte en PDF",
                                data=pdf_file,
                                file_name="Reporte_MultiHoja.pdf",
                                mime="application/pdf"
                            )
    except Exception as e:
        st.error(f"Error al procesar las hojas del Excel: {e}")
