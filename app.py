import io
import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
import openai

# Configuración de página
st.set_page_config(page_title="Generador de Reportes Corporativos", layout="centered")

st.title("📊 Generador Automatizado de Reportes")
st.write("Sube tu archivo Excel (.xlsx) para procesar los datos y descargar la presentación y el reporte en PDF.")

# 1. Configuración de API Key (Privacidad: la API de OpenAI no entrena con datos enviados vía API)
api_key = st.sidebar.text_input("OpenAI API Key", type="password", help="Tu API Key no se guarda en el servidor.")

# 2. Subida del archivo
uploaded_file = st.file_uploader("Arrastra tu archivo Excel aquí", type=["xlsx", "xls"])

def analizar_con_ia(resumen_texto, key):
    """Envía únicamente el resumen agregado a la API para mantener tokens bajos y privacidad alta."""
    client = openai.OpenAI(api_key=key)
    prompt = f"""
    Eres un analista de negocios senior. Analiza los siguientes datos resumidos y redacta:
    1. Tres conclusiones clave (bullet points concisos).
    2. Dos recomendaciones estratégicas de acción.
    
    Datos:
    {resumen_texto}
    
    Responde en formato directo, formal y conciso.
    """
    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[{"role": "user", "content": prompt}],
        temperature=0.3,
        max_tokens=400
    )
    return response.choices[0].message.content

def crear_powerpoint(df_resumen, analisis_ia):
    prs = Presentation()
    
    # Diapositiva 1: Portada
    slide_layout = prs.slide_layouts[0] # Título y subtítulo
    slide = prs.slides.add_slide(slide_layout)
    slide.shapes.title.text = "Informe Ejecutivo de Resultados"
    slide.placeholders[1].text = "Generado automáticamente vía Streamlit\\nConfidencial"
    
    # Diapositiva 2: Resumen Cuantitativo
    slide_layout = prs.slide_layouts[5] # Solo título
    slide = prs.slides.add_slide(slide_layout)
    slide.shapes.title.text = "Resumen de Métricas Clave"
    
    # Insertar tabla en PPTX
    rows, cols = min(len(df_resumen) + 1, 8), len(df_resumen.columns)
    table_shape = slide.shapes.add_table(rows, cols, Inches(1), Inches(2), Inches(8), Inches(3.5))
    table = table_shape.table
    
    # Encabezados
    for i, col_name in enumerate(df_resumen.columns):
        table.cell(0, i).text = str(col_name)
    
    # Filas
    for row_idx, row in df_resumen.head(7).iterrows():
        for col_idx, value in enumerate(row):
            table.cell(row_idx + 1, col_idx).text = str(value)
            
    # Diapositiva 3: Conclusiones IA
    slide = prs.slides.add_slide(slide_layout)
    slide.shapes.title.text = "Conclusiones y Recomendaciones"
    tx_box = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(4))
    tf = tx_box.text_frame
    tf.word_wrap = True
    tf.text = analisis_ia
    
    pptx_io = io.BytesIO()
    prs.save(pptx_io)
    pptx_io.seek(0)
    return pptx_io

def crear_pdf(df_resumen, analisis_ia):
    pdf_io = io.BytesIO()
    doc = SimpleDocTemplate(pdf_io, pagesize=letter, rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=40)
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle(name='TitleStyle', parent=styles['Heading1'], fontSize=18, textColor=colors.HexColor('#1E293B'))
    h2_style = ParagraphStyle(name='H2Style', parent=styles['Heading2'], fontSize=14, textColor=colors.HexColor('#2563EB'))
    body_style = ParagraphStyle(name='BodyStyle', parent=styles['Normal'], fontSize=10, leading=14)
    
    story = []
    story.append(Paragraph("Informe Ejecutivo de Desempeño", title_style))
    story.append(Spacer(1, 15))
    
    story.append(Paragraph("1. Análisis Estratégico y Conclusiones", h2_style))
    story.append(Spacer(1, 8))
    for parrafo in analisis_ia.split("\n"):
        if parrafo.strip():
            story.append(Paragraph(parrafo, body_style))
            story.append(Spacer(1, 4))
    story.append(Spacer(1, 15))
    
    story.append(Paragraph("2. Tabla de Datos Procesados", h2_style))
    story.append(Spacer(1, 8))
    
    # Preparar datos de la tabla
    tabla_data = [df_resumen.columns.tolist()] + df_resumen.head(10).values.tolist()
    # Convertir a texto para ReportLab
    tabla_data_str = [[str(cell) for cell in row] for row in tabla_data]
    
    t = Table(tabla_data_str)
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#F1F5F9')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.HexColor('#0F172A')),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 9),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CBD5E1')),
    ]))
    story.append(t)
    
    doc.build(story)
    pdf_io.seek(0)
    return pdf_io

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        st.success("Archivo cargado correctamente.")
        st.subheader("Vista previa de los datos")
        st.dataframe(df.head(5))
        
        # Validación de API Key
        if not api_key:
            st.warning("Ingresa tu OpenAI API Key en la barra lateral para generar el análisis con IA.")
        else:
            if st.button("🚀 Generar Reportes (.pptx y .pdf)"):
                with st.spinner("Procesando datos y redactando informe..."):
                    # Resumen numérico rápido para enviar a la IA
                    resumen_texto = df.describe().to_string()
                    
                    # Llamada a IA
                    analisis = analizar_con_ia(resumen_texto, api_key)
                    
                    # Generar archivos binarios en memoria
                    pptx_file = crear_powerpoint(df, analisis)
                    pdf_file = crear_pdf(df, analisis)
                    
                    st.success("¡Reportes generados con éxito!")
                    
                    col1, col2 = st.columns(2)
                    with col1:
                        st.download_button(
                            label="📥 Descargar PowerPoint (.pptx)",
                            data=pptx_file,
                            file_name="Reporte_Ejecutivo.pptx",
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                        )
                    with col2:
                        st.download_button(
                            label="📥 Descargar Reporte en PDF",
                            data=pdf_file,
                            file_name="Reporte_Ejecutivo.pdf",
                            mime="application/pdf"
                        )
    except Exception as e:
        st.error(f"Error al procesar el archivo: {e}")
