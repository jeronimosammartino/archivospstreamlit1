import io
import streamlit as st
import pandas as pd
import openpyxl
from pptx import Presentation
from pptx.util import Inches, Pt
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
import openai

st.set_page_config(page_title="Control de Métodos y Proyectos", layout="wide")

st.title("📊 Control de Métodos, HH y Costos de Proyecto")
st.write("Herramienta de análisis y generación de reportes ejecutivos para matrices de ingeniería y presupuestos.")

# Barra lateral
api_key = st.sidebar.text_input("OpenAI API Key", type="password", help="Tu API Key se usa únicamente en tu sesión.")
st.sidebar.markdown("---")

uploaded_file = st.file_uploader("Sube tu archivo Excel (.xlsx)", type=["xlsx", "xls"])

def leer_matriz_exacta(file_bytes, sheet_name):
    """
    Lee la hoja desenrollando celdas combinadas y preservando todas las filas y columnas.
    """
    wb = openpyxl.load_workbook(file_bytes, data_only=True)
    ws = wb[sheet_name]
    
    # 1. Propagar valores de celdas combinadas a todas las celdas del bloque
    merged_ranges = list(ws.merged_cells.ranges)
    for mr in merged_ranges:
        val = ws.cell(row=mr.min_row, column=mr.min_col).value
        ws.unmerge_cells(range_string=str(mr))
        for r in range(mr.min_row, mr.max_row + 1):
            for c in range(mr.min_col, mr.max_col + 1):
                ws.cell(row=r, column=c).value = val
                
    # 2. Extraer datos fila por fila
    filas = list(ws.iter_rows(values_only=True))
    if not filas:
        return pd.DataFrame()
        
    # 3. Detectar fila de encabezados
    header_idx = 0
    for idx, row in enumerate(filas[:10]):
        textos = [str(c).strip().upper() for c in row if c is not None]
        if any("CATEGORIA" in t or "GDF" in t or "PESO" in t for t in textos):
            header_idx = idx
            break
            
    headers = [str(c).strip().replace("\n", " ") if c is not None else f"Col_{i}" for i, c in enumerate(filas[header_idx])]
    
    datos = filas[header_idx + 1:]
    df = pd.DataFrame(datos, columns=headers)
    
    # Eliminar únicamente filas que sean 100% None
    df = df.dropna(how="all").reset_index(drop=True)
    
    # Asegurar que la columna de Categoría quede poblada
    if "Categoria" in df.columns:
        df["Categoria"] = df["Categoria"].ffill()
        
    return df

def analizar_con_ia(df, key):
    client = openai.OpenAI(api_key=key)
    
    # Extraer las filas de totales clave
    df_totales = df[df.apply(lambda r: r.astype(str).str.contains("Total", case=False).any(), axis=1)]
    resumen_totales = df_totales.to_string() if not df_totales.empty else df.to_string()
    
    prompt = f"""
    Eres el Ingeniero Jefe de Métodos, Procesos y Costos Industriales.
    Analiza la siguiente matriz consolidada de horas hombre (HH) y pesos del proyecto:
    
    {resumen_totales}
    
    DETALLE POR SUBCONJUNTO:
    {df[['Categoria', 'GDF', 'Peso', 'HH ESTRUCTURA', 'HH SOLDADURA', 'HH QA-QC']].to_string()}
    
    Genera un informe gerencial cuantitativo y riguroso estructurado en:
    1. Diagnóstico Ejecutivo: Variación global entre Oferta, Estimado de HdR, Estimado de OF y Real Total (destaca desvío total de peso y HH).
    2. Puestos Críticos y Sobrecostos: Disciplinas con mayor desvío (Estructura, Soldadura, Pintura, Mecanizado pesado/tornos, etc.).
    3. Productividad (Horas / Tonelada): Evolución del ratio HH/Tn entre lo ofertado y lo real ejecutado.
    4. Conclusiones y Plan de Acción Operativo.
    """
    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[{"role": "user", "content": prompt}],
        temperature=0.2,
        max_tokens=850
    )
    return response.choices[0].message.content

def crear_powerpoint(df, analisis_ia):
    prs = Presentation()
    
    # Slide 1: Portada
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = "Balance de Horas Hombre y Métodos"
    slide.placeholders[1].text = "Control de Desvíos: Oferta vs Estimado vs Real\nConfidencial"
    
    # Slide 2: Conclusiones IA
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = "Diagnóstico Estratégico y Desvíos"
    tx_box = slide.shapes.add_textbox(Inches(0.6), Inches(1.2), Inches(8.8), Inches(5.6))
    tf = tx_box.text_frame
    tf.word_wrap = True
    tf.text = analisis_ia
    
    # Slide 3: Matriz Comparativa de Totales
    df_totales = df[df.apply(lambda r: r.astype(str).str.contains("Total", case=False).any(), axis=1)]
    if not df_totales.empty:
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        slide.shapes.title.text = "Comparativa de Totales (Oferta vs Estimados vs Real)"
        
        cols_clave = [c for c in ['Categoria', 'Peso', 'HH ESTRUCTURA', 'HH SOLDADURA', 'HH AJUSTE', 'HH PINTURA', 'HH QA-QC'] if c in df.columns]
        df_sub = df_totales[cols_clave]
        
        rows, cols = len(df_sub) + 1, len(cols_clave)
        table_shape = slide.shapes.add_table(rows, cols, Inches(0.5), Inches(1.8), Inches(9.0), Inches(3.8))
        table = table_shape.table
        
        for i, col_name in enumerate(cols_clave):
            table.cell(0, i).text = str(col_name)
            
        for row_idx, (_, row) in enumerate(df_sub.iterrows()):
            for col_idx, col_name in enumerate(cols_clave):
                val = row[col_name]
                table.cell(row_idx + 1, col_idx).text = str(val) if pd.notna(val) else "-"
                
    pptx_io = io.BytesIO()
    prs.save(pptx_io)
    pptx_io.seek(0)
    return pptx_io

def crear_pdf(df, analisis_ia):
    pdf_io = io.BytesIO()
    doc = SimpleDocTemplate(pdf_io, pagesize=landscape(letter), rightMargin=25, leftMargin=25, topMargin=25, bottomMargin=25)
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle(name='TitleStyle', parent=styles['Heading1'], fontSize=14, textColor=colors.HexColor('#0F172A'))
    h2_style = ParagraphStyle(name='H2Style', parent=styles['Heading2'], fontSize=11, textColor=colors.HexColor('#1E40AF'))
    body_style = ParagraphStyle(name='BodyStyle', parent=styles['Normal'], fontSize=8.5, leading=12)
    
    story = []
    story.append(Paragraph("Informe Gerencial de Desvíos de Métodos y Horas Hombre", title_style))
    story.append(Spacer(1, 8))
    
    story.append(Paragraph("1. Evaluación y Dictamen Técnico", h2_style))
    story.append(Spacer(1, 4))
    for parrafo in analisis_ia.split("\n"):
        if parrafo.strip():
            story.append(Paragraph(parrafo, body_style))
            story.append(Spacer(1, 2))
    story.append(Spacer(1, 10))
    
    df_totales = df[df.apply(lambda r: r.astype(str).str.contains("Total", case=False).any(), axis=1)]
    if not df_totales.empty:
        story.append(Paragraph("2. Resumen Comparativo de Bloques y Totales", h2_style))
        story.append(Spacer(1, 4))
        
        cols_clave = [c for c in ['Categoria', 'Peso', 'HH ESTRUCTURA', 'HH SOLDADURA', 'HH AJUSTE', 'HH PINTURA', 'HH QA-QC'] if c in df.columns]
        tabla_data = [cols_clave]
        for _, row in df_totales[cols_clave].iterrows():
            tabla_data.append([str(row[c]) if pd.notna(row[c]) else "-" for c in cols_clave])
            
        t = Table(tabla_data)
        t.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#E2E8F0')),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#94A3B8')),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
        ]))
        story.append(t)
        
    doc.build(story)
    pdf_io.seek(0)
    return pdf_io

if uploaded_file is not None:
    try:
        file_bytes = io.BytesIO(uploaded_file.read())
        wb_check = openpyxl.load_workbook(file_bytes, read_only=True)
        nombres_hojas = wb_check.sheetnames
        wb_check.close()
        
        st.success(f"Archivo cargado. Hojas disponibles: {', '.join(nombres_hojas)}")
        
        hoja_activa = st.selectbox(
            "Selecciona la hoja a procesar:",
            options=nombres_hojas,
            index=1 if len(nombres_hojas) > 1 else 0
        )
        
        file_bytes.seek(0)
        df_completo = leer_matriz_exacta(file_bytes, hoja_activa)
        
        st.subheader(f"Vista Completa de '{hoja_activa}' ({df_completo.shape[0]} filas × {df_completo.shape[1]} columnas)")
        st.dataframe(df_completo, height=550)
        
        if not api_key:
            st.warning("Ingresa tu OpenAI API Key en la barra lateral para generar el informe y las diapositivas.")
        else:
            if st.button("🚀 Generar Informe y Presentación"):
                with st.spinner("Generando diagnóstico con IA y compilando archivos..."):
                    analisis = analizar_con_ia(df_completo, api_key)
                    pptx_file = crear_powerpoint(df_completo, analisis)
                    pdf_file = crear_pdf(df_completo, analisis)
                    
                    st.success("¡Informe generado con éxito!")
                    col1, col2 = st.columns(2)
                    with col1:
                        st.download_button(
                            label="📥 Descargar PowerPoint (.pptx)",
                            data=pptx_file,
                            file_name=f"Informe_{hoja_activa}.pptx",
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                        )
                    with col2:
                        st.download_button(
                            label="📥 Descargar PDF (Horizontal)",
                            data=pdf_file,
                            file_name=f"Informe_{hoja_activa}.pdf",
                            mime="application/pdf"
                        )
    except Exception as e:
        st.error(f"Error procesando la matriz: {e}")
