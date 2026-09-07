import io
import streamlit as st
import pandas as pd
import openpyxl
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(page_title="Control de Métodos y Horas Hombre", layout="wide")

st.title("📊 Control y Balances de Horas Hombre (HH) y Rendimientos")
st.write("Panel interactivo de control operacional: Oferta vs Estimado (HdR) vs Real.")

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

def procesar_datos(df):
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
    
    df_pesos = df_det[['Estado', 'GDF', 'Peso']].drop_duplicates()
    df_pesos['Peso'] = pd.to_numeric(df_pesos['Peso'], errors='coerce').fillna(0)
    
    return df_long, df_pesos

if uploaded_file:
    file_bytes = io.BytesIO(uploaded_file.read())
    wb_check = openpyxl.load_workbook(file_bytes, read_only=True)
    hojas = wb_check.sheetnames
    wb_check.close()
    
    hoja = st.sidebar.selectbox("Hoja activa:", hojas, index=1 if len(hojas) > 1 else 0)
    file_bytes.seek(0)
    df_raw = leer_matriz_exacta(file_bytes, hoja)
    
    df_long, df_pesos = procesar_datos(df_raw)
    
    # --- FILTROS INTERACTIVOS ---
    st.sidebar.markdown("---")
    st.sidebar.subheader("🔍 Filtros de Visualización")
    
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
        
    # --- CÁLCULOS PRINCIPALES ---
    tot_horas = df_f.groupby('Estado')['Horas'].sum().to_dict()
    h_ofe = tot_horas.get("OFERTADAS", 0)
    h_est = tot_horas.get("ESTIMADAS", 0)
    h_real = tot_horas.get("REALES", 0)
    
    tot_peso = df_p_f.groupby('Estado')['Peso'].sum().to_dict()
    p_ofe = tot_peso.get("OFERTADAS", 0)
    p_est = tot_peso.get("ESTIMADAS", 0)
    p_real = tot_peso.get("REALES", 0)
    
    # Rendimientos (Peso / HH)
    rend_ofe = (p_ofe / h_ofe) if h_ofe > 0 else 0
    rend_est = (p_est / h_est) if h_est > 0 else 0
    rend_real = (p_real / h_real) if h_real > 0 else 0
    
    color_map = {"OFERTADAS": "#F59E0B", "ESTIMADAS": "#06B6D4", "REALES": "#6366F1"}
    color_map_rend = {"OFERTADO": "#F59E0B", "ESTIMADO": "#06B6D4", "REAL": "#6366F1"}
    orden_estados = ["OFERTADAS", "ESTIMADAS", "REALES"]
    
    # --- PESTAÑAS (PÁGINAS DE POWER BI) ---
    t1, t2, t3, t4, t5, t6 = st.tabs([
        "HORAS 1",
        "HORAS 2",
        "ESTIMADAS VS OFERTADAS",
        "REALES VS ESTIMADAS",
        "REALES VS OFERTADAS",
        "RENDIMIENTOS"
    ])
    
    # 1. HORAS 1
    with t1:
        st.subheader("HORAS")
        df_bar_global = pd.DataFrame({"Estado": orden_estados, "Horas": [h_ofe, h_est, h_real]})
        fig_h1 = px.bar(
            df_bar_global, x="Estado", y="Horas", color="Estado",
            text_auto=",.0f", color_discrete_map=color_map
        )
        fig_h1.update_layout(height=400, showlegend=False, xaxis_title="", yaxis_title="Horas")
        st.plotly_chart(fig_h1, use_container_width=True)
        
        c1, c2, c3 = st.columns(3)
        v_eo = h_est - h_ofe
        p_eo = (v_eo / h_ofe * 100) if h_ofe else 0
        c1.metric("VARIACION HH ESTIMADO VS OFERTADO", f"{int(v_eo):,}", f"{p_eo:.2f}%")
        
        v_re = h_real - h_est
        p_re = (v_re / h_est * 100) if h_est else 0
        c2.metric("VARIACION HH REAL VS ESTIMADO", f"{int(v_re):,}", f"{p_re:.2f}%")
        
        v_ro = h_real - h_ofe
        p_ro = (v_ro / h_ofe * 100) if h_ofe else 0
        c3.metric("VARIACION HH REALES VS OFERTADAS", f"{int(v_ro):,}", f"{p_ro:.2f}%")
        
    # 2. HORAS 2
    with t2:
        st.subheader("HORAS POR TAREA")
        df_tarea = df_f.groupby(['Recurso', 'Estado'])['Horas'].sum().reset_index()
        fig_t = px.bar(
            df_tarea, x="Recurso", y="Horas", color="Estado",
            barmode="group", category_orders={"Estado": orden_estados},
            color_discrete_map=color_map
        )
        fig_t.update_layout(height=420, xaxis_tickangle=-45, yaxis_title="", legend_title="")
        st.plotly_chart(fig_t, use_container_width=True)
        
        st.subheader("HORAS POR GDF")
        df_gdf = df_f.groupby(['GDF', 'Estado'])['Horas'].sum().reset_index()
        fig_g = px.bar(
            df_gdf, x="GDF", y="Horas", color="Estado",
            barmode="group", category_orders={"Estado": orden_estados},
            color_discrete_map=color_map
        )
        fig_g.update_layout(height=420, yaxis_title="", legend_title="")
        st.plotly_chart(fig_g, use_container_width=True)
        
    def crear_cascada(df_base, col_agrupacion, estado_fin, estado_ini, titulo):
        piv = df_base.pivot_table(index=col_agrupacion, columns='Estado', values='Horas', aggfunc='sum').fillna(0)
        if estado_fin in piv.columns and estado_ini in piv.columns:
            piv['Desvio'] = piv[estado_fin] - piv[estado_ini]
            piv = piv.sort_values(by='Desvio', ascending=False).reset_index()
            
            fig = go.Figure(go.Waterfall(
                orientation="v",
                measure=["relative"] * len(piv) + ["total"],
                x=piv[col_agrupacion].tolist() + ["Total"],
                y=piv['Desvio'].tolist() + [piv['Desvio'].sum()],
                textposition="outside",
                decreasing={"marker": {"color": "#10B981"}},
                increasing={"marker": {"color": "#EF4444"}},
                totals={"marker": {"color": "#06B6D4"}}
            ))
            fig.update_layout(title=titulo, height=400, xaxis_tickangle=-45, yaxis_title="")
            return fig
        return None

    # 3. ESTIMADAS VS OFERTADAS
    with t3:
        f_g_eo = crear_cascada(df_f, 'GDF', 'ESTIMADAS', 'OFERTADAS', "HH ESTIMADO VS OFERTADO POR GDF")
        if f_g_eo: st.plotly_chart(f_g_eo, use_container_width=True)
        
        f_t_eo = crear_cascada(df_f, 'Recurso', 'ESTIMADAS', 'OFERTADAS', "HH ESTIMADO VS OFERTADO POR TAREA")
        if f_t_eo: st.plotly_chart(f_t_eo, use_container_width=True)

    # 4. REALES VS ESTIMADAS
    with t4:
        f_g_re = crear_cascada(df_f, 'GDF', 'REALES', 'ESTIMADAS', "HH REALES VS ESTIMADAS POR GDF")
        if f_g_re: st.plotly_chart(f_g_re, use_container_width=True)
        
        f_t_re = crear_cascada(df_f, 'Recurso', 'REALES', 'ESTIMADAS', "HH REALES VS ESTIMADAS POR TAREA")
        if f_t_re: st.plotly_chart(f_t_re, use_container_width=True)

    # 5. REALES VS OFERTADAS
    with t5:
        f_g_ro = crear_cascada(df_f, 'GDF', 'REALES', 'OFERTADAS', "HH REALES VS OFERTADAS POR GDF")
        if f_g_ro: st.plotly_chart(f_g_ro, use_container_width=True)
        
        f_t_ro = crear_cascada(df_f, 'Recurso', 'REALES', 'OFERTADAS', "VARIACION HH REALES VS OFERTADAS by RECURSO")
        if f_t_ro: st.plotly_chart(f_t_ro, use_container_width=True)

    # 6. RENDIMIENTOS
    with t6:
        st.subheader("RENDIMIENTOS (Peso / HH)")
        df_rend = pd.DataFrame({
            "Estado": ["OFERTADO", "ESTIMADO", "REAL"],
            "Rendimiento": [rend_ofe, rend_est, rend_real]
        })
        fig_r = px.bar(
            df_rend, x="Estado", y="Rendimiento", color="Estado",
            text=df_rend["Rendimiento"].apply(lambda v: f"{v:.2f}"),
            color_discrete_map=color_map_rend
        )
        fig_r.update_layout(height=400, showlegend=False, xaxis_title="", yaxis_title="Kg / HH")
        st.plotly_chart(fig_r, use_container_width=True)
        
        cr1, cr2, cr3 = st.columns(3)
        
        d_rend_eo = rend_est - rend_ofe
        p_rend_eo = (d_rend_eo / rend_ofe * 100) if rend_ofe else 0
        cr1.metric("REND ESTIMADO VS OFERTADO", f"{d_rend_eo:.2f}", f"{p_rend_eo:.2f}% (%V E-O)")
        
        d_rend_re = rend_real - rend_est
        p_rend_re = (d_rend_re / rend_est * 100) if rend_est else 0
        cr2.metric("REND REAL VS ESTIM", f"{d_rend_re:.2f}", f"{p_rend_re:.2f}% (%V R-E)")
        
        d_rend_ro = rend_real - rend_ofe
        p_rend_ro = (d_rend_ro / rend_ofe * 100) if rend_ofe else 0
        cr3.metric("REND REAL VS OFERTADO", f"{d_rend_ro:.2f}", f"{p_rend_ro:.2f}% (%V R-O)")
