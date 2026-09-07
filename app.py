import io
import streamlit as st
import pandas as pd
import openpyxl
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

st.set_page_config(page_title="Control de HH - Métodos", layout="wide")

st.title("📊 Control de Horas Hombre (HH): Métodos y Operaciones")
st.write("Generador de balances y reportes ejecutivos (Oferta vs Estimado vs Real).")

uploaded_file = st.file_uploader("Sube la matriz del proyecto (.xlsx)", type=["xlsx", "xls"])

def leer_matriz_exacta(file_bytes, sheet_name):
    wb = openpyxl.load_workbook(file_bytes, data_only=True)
    ws = wb[sheet_name]
    
    # Resolver celdas combinadas
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

def calcular_metricas(df):
    cols_hh = [c for c in df.columns if c.startswith("HH ") or c in ["Autycontrol", "MET"]]
    
    # 1. Totales Generales
    fila_ofe = df[df['Categoria'].astype(str).str.contains("Total Oferta", case=False, na=False)]
    fila_est = df[df['Categoria'].astype(str).str.contains("Total Estimado", case=False, na=False) & 
                  ~df['Categoria'].astype(str).str.contains("OF", case=False, na=False)]
    fila_real = df[df['Categoria'].astype(str).str.contains("Total Proyecto", case=False, na=False)]
    
    tot_ofe = fila_ofe[cols_hh].sum(axis=1).values[0] if not fila_ofe.empty else 0
    tot_est = fila_est[cols_hh].sum(axis=1).values[0] if not fila_est.empty else 0
    tot_real = fila_real[cols_hh].sum(axis=1).values[0] if not fila_real.empty else 0
    
    # Variaciones
    var_est_ofe = tot_est - tot_ofe
    pct_est_ofe = (var_est_ofe / tot_ofe * 100) if tot_ofe else 0
    
    var_real_est = tot_real - tot_est
    pct_real_est = (var_real_est / tot_est * 100) if tot_est else 0
    
    var_real_ofe = tot_real - tot_ofe
    pct_real_ofe = (var_real_ofe / tot_ofe * 100) if tot_ofe else 0
    
    # Desvíos por tarea (Real vs Oferta)
    s_ofe = fila_ofe[cols_hh].iloc[0] if not fila_ofe.empty else pd.Series()
    s_real = fila_real[cols_hh].iloc[0] if not fila_real.empty else pd.Series()
    desvios_tarea = (s_real - s_ofe).sort_values(ascending=False)
    
    return {
        "ofe": tot_ofe, "est": tot_est, "real": tot_real,
        "var_est_ofe": var_est_ofe, "pct_est_ofe": pct_est_ofe,
        "var_real_est": var_real_est, "pct_real_est": pct_real_est,
        "var_real_ofe": var_real_ofe, "pct_real_ofe": pct_real_ofe,
        "desvios_tarea": desvios_tarea
    }

def generar_pptx(m):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = "Informe de Desvíos de Horas Hombre"
    slide.placeholders[1].text = f"Variación Total: +{m['var_real_ofe']:,} HH ({m['pct_real_ofe']:.2f}%)\nIMPSA - Métodos"
    
    # Diapositiva de KPIs
    slide2 = prs.slides.add_slide(prs.slide_layouts[5])
    slide2.shapes.title.text = "Métricas Globales de Horas Hombre"
    
    tb = slide2.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(4))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.text = f"• Total Ofertadas: {m['ofe']:,} HH\n• Total Estimadas (HdR): {m['est']:,} HH\n• Total Reales: {m['real']:,} HH\n\n"
    tf.text += f"1. Estimado vs Ofertado: +{m['var_est_ofe']:,} HH (+{m['pct_est_ofe']:.2f}%)\n"
    tf.text += f"2. Real vs Estimado: +{m['var_real_est']:,} HH (+{m['pct_real_est']:.2f}%)\n"
    tf.text += f"3. Real vs Ofertado: +{m['var_real_ofe']:,} HH (+{m['pct_real_ofe']:.2f}%)\n\n"
    tf.text += f"Top Puestos Críticos con Mayor Sobrecosto:\n"
    for tarea, val in m['desvios_tarea'].head(5).items():
        tf.text += f"  - {tarea}: +{val:,} HH\n"
        
    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out

if uploaded_file:
    file_bytes = io.BytesIO(uploaded_file.read())
    wb_check = openpyxl.load_workbook(file_bytes, read_only=True)
    hojas = wb_check.sheetnames
    wb_check.close()
    
    hoja = st.selectbox("Hoja a procesar:", hojas, index=1 if len(hojas) > 1 else 0)
    file_bytes.seek(0)
    df = leer_matriz_exacta(file_bytes, hoja)
    
    st.subheader(f"Vista de '{hoja}'")
    st.dataframe(df, height=350)
    
    m = calcular_metricas(df)
    
    st.markdown("---")
    st.subheader("Indicadores Clave de Desvío (Idéntico a Power BI)")
    
    c1, c2, c3 = st.columns(3)
    c1.metric("Variación Estimado vs Ofertado", f"{m['var_est_ofe']:,} HH", f"{m['pct_est_ofe']:.2f}%")
    c2.metric("Variación Real vs Estimado", f"{m['var_real_est']:,} HH", f"{m['pct_real_est']:.2f}%")
    c3.metric("Variación Real vs Ofertado", f"{m['var_real_ofe']:,} HH", f"{m['pct_real_ofe']:.2f}%")
    
    st.markdown("---")
    st.write("### Top Desvíos por Tarea (Sobrecostos)")
    st.dataframe(m['desvios_tarea'].head(8).to_frame(name="Diferencia HH (Real - Oferta)"))
    
    pptx_data = generar_pptx(m)
    st.download_button("📥 Descargar Presentación (.pptx)", pptx_data, "Reporte_Horas_Hombre.pptx", "application/vnd.openxmlformats-officedocument.presentationml.presentation")
