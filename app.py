import io
import streamlit as st
import pandas as pd
import openpyxl
import plotly.express as px
import plotly.graph_objects as go
from pptx import Presentation
from pptx.util import Inches

st.set_page_config(page_title="Dashboard Interactivo de Horas Hombre", layout="wide")

st.title("📊 Control y Balances de Horas Hombre (HH)")
st.write("Panel interactivo de control: Oferta vs Estimado (HdR) vs Real con filtros dinámicos.")

uploaded_file = st.file_uploader("Sube la matriz del proyecto (.xlsx)", type=["xlsx", "xls"])

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

def transformar_a_base_larga(df):
    cols_recurso = [c for c in df.columns if c.startswith("HH ") or c in ["Autycontrol", "MET"]]
    df_detalle = df[~df['GDF'].astype(str).str.contains("Total", case=False, na=False)].copy()
    
    def normalizar_cat(val):
        v = str(val).upper()
        if "OFERTA" in v:
            return "OFERTADAS"
        elif "HDR" in v:
            return "ESTIMADAS"
        elif "REAL" in v:
            return "REALES"
        return None

    df_detalle['Estado'] = df_detalle['Categoria'].apply(normalizar_cat)
    df_detalle = df_detalle.dropna(subset=['Estado'])
    
    df_long = df_detalle.melt(
        id_vars=['Estado', 'GDF'],
        value_vars=cols_recurso,
        var_name='Recurso',
        value_name='Horas'
    )
    df_long['Horas'] = pd.to_numeric(df_long['Horas'], errors='coerce').fillna(0)
    return df_long

if uploaded_file:
    file_bytes = io.BytesIO(uploaded_file.read())
    wb_check = openpyxl.load_workbook(file_bytes, read_only=True)
    hojas = wb_check.sheetnames
    wb_check.close()
    
    hoja = st.sidebar.selectbox("Hoja activa:", hojas, index=1 if len(hojas) > 1 else 0)
    file_bytes.seek(0)
    df_raw = leer_matriz_exacta(file_bytes, hoja)
    
    df_long = transformar_a_base_larga(df_raw)
    
    # --- FILTROS INTERACTIVOS ---
    st.sidebar.markdown("---")
    st.sidebar.subheader("🔍 Filtros Dinámicos")
    
    lista_recursos = sorted(df_long['Recurso'].unique().tolist())
    filtro_recurso = st.sidebar.multiselect("Filtrar por RECURSO / TAREA:", options=lista_recursos, default=[])
    
    lista_gdf = sorted(df_long['GDF'].dropna().unique().tolist())
    filtro_gdf = st.sidebar.multiselect("Filtrar por GDF (Subconjunto):", options=lista_gdf, default=[])
    
    df_filtrado = df_long.copy()
    if filtro_recurso:
        df_filtrado = df_filtrado[df_filtrado['Recurso'].isin(filtro_recurso)]
    if filtro_gdf:
        df_filtrado = df_filtrado[df_filtrado['GDF'].isin(filtro_gdf)]
        
    # --- KPIS ---
    totales_estado = df_filtrado.groupby('Estado')['Horas'].sum().to_dict()
    tot_ofe = totales_estado.get("OFERTADAS", 0)
    tot_est = totales_estado.get("ESTIMADAS", 0)
    tot_real = totales_estado.get("REALES", 0)
    
    var_est_ofe = tot_est - tot_ofe
    pct_est_ofe = (var_est_ofe / tot_ofe * 100) if tot_ofe else 0
    
    var_real_est = tot_real - tot_est
    pct_real_est = (var_real_est / tot_est * 100) if tot_est else 0
    
    var_real_ofe = tot_real - tot_ofe
    pct_real_ofe = (var_real_ofe / tot_ofe * 100) if tot_ofe else 0

    c1, c2, c3 = st.columns(3)
    c1.metric("VARIACIÓN HH ESTIMADO VS OFERTADO", f"{int(var_est_ofe):,} HH", f"{pct_est_ofe:.2f}%")
    c2.metric("VARIACIÓN HH REAL VS ESTIMADO", f"{int(var_real_est):,} HH", f"{pct_real_est:.2f}%")
    c3.metric("VARIACIÓN HH REALES VS OFERTADAS", f"{int(var_real_ofe):,} HH", f"{pct_real_ofe:.2f}%")
    
    st.markdown("---")
    
    # --- GRÁFICOS INTERACTIVOS ---
    color_map = {
        "OFERTADAS": "#F59E0B",
        "ESTIMADAS": "#06B6D4",
        "REALES": "#6366F1"
    }
    orden_estados = ["OFERTADAS", "ESTIMADAS", "REALES"]

    st.subheader("📊 Comparativa Global de Horas")
    df_tot_plot = pd.DataFrame({
        "Estado": orden_estados,
        "Horas": [tot_ofe, tot_est, tot_real]
    })
    fig_global = px.bar(
        df_tot_plot, x="Estado", y="Horas", color="Estado",
        text_auto=",.0f",
        color_discrete_map=color_map,
        title="Total de Horas Hombre por Estado"
    )
    fig_global.update_layout(height=380, showlegend=False, xaxis_title="", yaxis_title="Horas Hombre (HH)")
    st.plotly_chart(fig_global, use_container_width=True)
    
    col_g1, col_g2 = st.columns(2)
    
    with col_g1:
        st.subheader("📌 Horas por Tarea")
        df_tarea = df_filtrado.groupby(['Recurso', 'Estado'])['Horas'].sum().reset_index()
        fig_tarea = px.bar(
            df_tarea, x="Recurso", y="Horas", color="Estado",
            barmode="group",
            category_orders={"Estado": orden_estados},
            color_discrete_map=color_map
        )
        fig_tarea.update_layout(height=450, xaxis_tickangle=-45, yaxis_title="HH", legend_title="")
        st.plotly_chart(fig_tarea, use_container_width=True)
        
    with col_g2:
        st.subheader("🏭 Horas por GDF (Subconjunto)")
        df_gdf = df_filtrado.groupby(['GDF', 'Estado'])['Horas'].sum().reset_index()
        fig_gdf = px.bar(
            df_gdf, x="GDF", y="Horas", color="Estado",
            barmode="group",
            category_orders={"Estado": orden_estados},
            color_discrete_map=color_map
        )
        fig_gdf.update_layout(height=450, yaxis_title="HH", legend_title="")
        st.plotly_chart(fig_gdf, use_container_width=True)
        
    st.subheader("🌊 Variación de Horas (Cascada por Tarea: Real vs Oferta)")
    df_pivot_tarea = df_filtrado.pivot_table(index='Recurso', columns='Estado', values='Horas', aggfunc='sum').fillna(0)
    if 'OFERTADAS' in df_pivot_tarea.columns and 'REALES' in df_pivot_tarea.columns:
        df_pivot_tarea['Desvío'] = df_pivot_tarea['REALES'] - df_pivot_tarea['OFERTADAS']
        df_desv = df_pivot_tarea.sort_values(by='Desvío', ascending=False).reset_index()
        
        fig_waterfall = go.Figure(go.Waterfall(
            name="Desvío",
            orientation="v",
            measure=["relative"] * len(df_desv) + ["total"],
            x=df_desv['Recurso'].tolist() + ["Total"],
            y=df_desv['Desvío'].tolist() + [df_desv['Desvío'].sum()],
            textposition="outside",
            decreasing={"marker": {"color": "#10B981"}},
            increasing={"marker": {"color": "#EF4444"}},
            totals={"marker": {"color": "#06B6D4"}}
        ))
        fig_waterfall.update_layout(height=450, xaxis_tickangle=-45, yaxis_title="Diferencia HH (Real - Oferta)")
        st.plotly_chart(fig_waterfall, use_container_width=True)    
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
