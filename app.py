import io
import streamlit as st
import pandas as pd
import openpyxl
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(page_title="Control de Métodos, Costos y HH", layout="wide")

st.title("📊 Control de Métodos: Horas Hombre, Costos y Rendimientos")
st.write("Tablero integral de control operacional (100% equivalente a Power BI).")

uploaded_file = st.file_uploader("Sube la matriz del proyecto (.xlsx)", type=["xlsx", "xls"])

def leer_fabricost(file_bytes):
    try:
        df_costos = pd.read_excel(file_bytes, sheet_name="FABRICOST")
        col_recurso = None
        col_tarifa = None
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
    
    color_map = {"OFERTADAS": "#F59E0B", "ESTIMADAS": "#06B6D4", "REALES": "#6366F1"}
    orden_estados = ["OFERTADAS", "ESTIMADAS", "REALES"]

    def formato_m_k(val):
        if abs(val) >= 1_000_000:
            return f"{val/1_000_000:.2f}M"
        elif abs(val) >= 1_000:
            return f"{val/1_000:.2f}K"
        return f"{val:.2f}"

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

    # --- LAS 11 PESTAÑAS COMPLETAS ---
    tabs = st.tabs([
        "HORAS 1", "HORAS 2", "ESTIMADAS VS OFERTADAS", "REALES VS ESTIMADAS", "REALES VS OFERTADAS",
        "COSTOS", "COSTOS 2", "COSTOS ESTIM VS OFERT", "COSTOS REAL VS ESTIM", "COSTOS REAL VS OF",
        "RENDIMIENTOS"
    ])
    
    # 1. HORAS 1
    with tabs[0]:
        st.subheader("HORAS")
        df_bh = pd.DataFrame({"Estado": orden_estados, "Horas": [h_ofe, h_est, h_real]})
        fig_h = px.bar(df_bh, x="Estado", y="Horas", color="Estado", text_auto=",.0f", color_discrete_map=color_map)
        fig_h.update_layout(height=400, showlegend=False, xaxis_title="", yaxis_title="Horas")
        st.plotly_chart(fig_h, use_container_width=True)
        
        c1, c2, c3 = st.columns(3)
        v_eo = h_est - h_ofe
        c1.metric("VARIACION HH ESTIMADO VS OFERTADO", f"{int(v_eo):,}", f"{(v_eo/h_ofe*100):.2f}%" if h_ofe else "0%")
        v_re = h_real - h_est
        c2.metric("VARIACION HH REAL VS ESTIMADO", f"{int(v_re):,}", f"{(v_re/h_est*100):.2f}%" if h_est else "0%")
        v_ro = h_real - h_ofe
        c3.metric("VARIACION HH REALES VS OFERTADAS", f"{int(v_ro):,}", f"{(v_ro/h_ofe*100):.2f}%" if h_ofe else "0%")

    # 2. HORAS 2
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

    # 3. ESTIMADAS VS OFERTADAS (HH)
    with tabs[2]:
        fig1 = crear_cascada(df_f, 'GDF', 'Horas', 'ESTIMADAS', 'OFERTADAS', "HH ESTIMADO VS OFERTADO POR GDF")
        if fig1: st.plotly_chart(fig1, use_container_width=True)
        fig2 = crear_cascada(df_f, 'Recurso', 'Horas', 'ESTIMADAS', 'OFERTADAS', "HH ESTIMADO VS OFERTADO POR TAREA")
        if fig2: st.plotly_chart(fig2, use_container_width=True)

    # 4. REALES VS ESTIMADAS (HH)
    with tabs[3]:
        fig1 = crear_cascada(df_f, 'GDF', 'Horas', 'REALES', 'ESTIMADAS', "HH REALES VS ESTIMADAS POR GDF")
        if fig1: st.plotly_chart(fig1, use_container_width=True)
        fig2 = crear_cascada(df_f, 'Recurso', 'Horas', 'REALES', 'ESTIMADAS', "HH REALES VS ESTIMADAS POR TAREA")
        if fig2: st.plotly_chart(fig2, use_container_width=True)

    # 5. REALES VS OFERTADAS (HH)
    with tabs[4]:
        fig1 = crear_cascada(df_f, 'GDF', 'Horas', 'REALES', 'OFERTADAS', "HH REALES VS OFERTADAS POR GDF")
        if fig1: st.plotly_chart(fig1, use_container_width=True)
        fig2 = crear_cascada(df_f, 'Recurso', 'Horas', 'REALES', 'OFERTADAS', "VARIACION HH REALES VS OFERTADAS by RECURSO")
        if fig2: st.plotly_chart(fig2, use_container_width=True)

    # 6. COSTOS
    with tabs[5]:
        st.subheader("COSTOS")
        df_bc = pd.DataFrame({"Estado": orden_estados, "Costo": [c_ofe, c_est, c_real]})
        df_bc["Etiqueta"] = df_bc["Costo"].apply(formato_m_k)
        fig_c = px.bar(df_bc, x="Estado", y="Costo", color="Estado", text="Etiqueta", color_discrete_map=color_map)
        fig_c.update_layout(height=400, showlegend=False, xaxis_title="", yaxis_title="Costo Total")
        st.plotly_chart(fig_c, use_container_width=True)
        
        c1, c2, c3 = st.columns(3)
        vc_eo = c_est - c_ofe
        c1.metric("COSTOS ESTIMADOS VS OFERTADOS", formato_m_k(vc_eo), f"{(vc_eo/c_ofe*100):.2f}%" if c_ofe else "0%")
        vc_re = c_real - c_est
        c2.metric("COSTOS REALES VS ESTIMADOS", formato_m_k(vc_re), f"{(vc_re/c_est*100):.2f}%" if c_est else "0%")
        vc_ro = c_real - c_ofe
        c3.metric("COSTOS REALES VS OFERTADOS", formato_m_k(vc_ro), f"{(vc_ro/c_ofe*100):.2f}%" if c_ofe else "0%")

    # 7. COSTOS 2
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

    # 8. COSTOS ESTIM VS OFERT
    with tabs[7]:
        fig1 = crear_cascada(df_f, 'GDF', 'Costo', 'ESTIMADAS', 'OFERTADAS', "COSTOS ESTIMADOS VS OFERTADOS POR GDF")
        if fig1: st.plotly_chart(fig1, use_container_width=True)
        fig2 = crear_cascada(df_f, 'Recurso', 'Costo', 'ESTIMADAS', 'OFERTADAS', "COSTOS ESTIMADOS VS OFERTADOS POR TAREA")
        if fig2: st.plotly_chart(fig2, use_container_width=True)

    # 9. COSTOS REAL VS ESTIM
    with tabs[8]:
        fig1 = crear_cascada(df_f, 'GDF', 'Costo', 'REALES', 'ESTIMADAS', "COSTOS REALES VS ESTIMADOS POR GDF")
        if fig1: st.plotly_chart(fig1, use_container_width=True)
        fig2 = crear_cascada(df_f, 'Recurso', 'Costo', 'REALES', 'ESTIMADAS', "COSTOS REALES VS ESTIMADOS POR RECURSO")
        if fig2: st.plotly_chart(fig2, use_container_width=True)

    # 10. COSTOS REAL VS OF
    with tabs[9]:
        fig1 = crear_cascada(df_f, 'GDF', 'Costo', 'REALES', 'OFERTADAS', "COSTOS REALES VS OFERTADOS POR GDF")
        if fig1: st.plotly_chart(fig1, use_container_width=True)
        fig2 = crear_cascada(df_f, 'Recurso', 'Costo', 'REALES', 'OFERTADAS', "COSTOS REALES VS OFERTADOS POR RECURSO")
        if fig2: st.plotly_chart(fig2, use_container_width=True)

    # 11. RENDIMIENTOS
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
        d_eo = rend_est - rend_ofe
        cr1.metric("REND ESTIMADO VS OFERTADO", f"{d_eo:.2f}", f"{(d_eo/rend_ofe*100):.2f}% (%V E-O)" if rend_ofe else "0%")
        d_re = rend_real - rend_est
        cr2.metric("REND REAL VS ESTIM", f"{d_re:.2f}", f"{(d_re/rend_est*100):.2f}% (%V R-E)" if rend_est else "0%")
        d_ro = rend_real - rend_ofe
        cr3.metric("REND REAL VS OFERTADO", f"{d_ro:.2f}", f"{(d_ro/rend_ofe*100):.2f}% (%V R-O)" if rend_ofe else "0%")
