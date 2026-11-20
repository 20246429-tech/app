import os
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
import streamlit as st

st.set_page_config(
    page_title="Dashboard Prescriptivo - Retail",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- CSS PREMIUM ---
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    .main { background-color: #0a0a0f; font-family: 'Inter', sans-serif; }

    /* KPI Cards */
    .kpi-row { display: flex; gap: 12px; margin-bottom: 20px; }
    .kpi {
        flex: 1;
        background: linear-gradient(145deg, #12121a, #1a1a2e);
        border: 1px solid rgba(255,255,255,0.06);
        border-radius: 16px;
        padding: 18px 14px;
        text-align: center;
    }
    .kpi-label { font-size: 10px; color: #6b7280; font-weight: 600; text-transform: uppercase; letter-spacing: 1.2px; margin-bottom: 8px; }
    .kpi-val { font-size: 26px; font-weight: 700; letter-spacing: -0.5px; }
    .c-green { color: #34d399; }
    .c-blue { color: #60a5fa; }
    .c-purple { color: #a78bfa; }
    .c-orange { color: #fb923c; }
    .c-red { color: #f87171; }
    .c-white { color: #e2e8f0; }

    /* Scenario Cards */
    .sc-grid { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 16px; margin: 16px 0; }
    .sc {
        background: linear-gradient(145deg, #12121a, #1a1a2e);
        border: 1px solid rgba(255,255,255,0.06);
        border-radius: 16px;
        padding: 24px 20px;
    }
    .sc-optimal { border: 2px solid #34d399; box-shadow: 0 0 30px rgba(52,211,153,0.08); }
    .sc-head { display: flex; align-items: center; justify-content: space-between; margin-bottom: 16px; }
    .sc-name { font-size: 16px; font-weight: 700; color: #e2e8f0; }
    .badge { padding: 4px 12px; border-radius: 20px; font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; }
    .badge-blue { background: rgba(96,165,250,0.15); color: #60a5fa; }
    .badge-green { background: rgba(52,211,153,0.15); color: #34d399; }
    .badge-red { background: rgba(248,113,113,0.15); color: #f87171; }
    .sc-row { display: flex; justify-content: space-between; padding: 8px 0; border-bottom: 1px solid rgba(255,255,255,0.04); }
    .sc-row:last-child { border-bottom: none; }
    .sc-metric { font-size: 11px; color: #6b7280; }
    .sc-val { font-size: 14px; font-weight: 600; color: #e2e8f0; }

    /* Recommendation */
    .rec-card {
        background: linear-gradient(145deg, #0a1a0f, #12201a);
        border: 1px solid rgba(52,211,153,0.2);
        border-radius: 16px;
        padding: 28px 24px;
        max-width: 700px;
    }
    .rec-title { font-size: 20px; font-weight: 700; color: #34d399; margin-bottom: 12px; }
    .rec-sub { font-size: 14px; color: #94a3b8; line-height: 1.6; }
    .rec-highlight { color: #34d399; font-weight: 600; }

    /* Sidebar */
    div[data-testid="stSidebar"] { background: #0f0f18; border-right: 1px solid rgba(255,255,255,0.04); }
    .sidebar-title { font-size: 20px; font-weight: 700; color: #e2e8f0; margin-bottom: 4px; }
    .sidebar-sub { font-size: 12px; color: #6b7280; margin-bottom: 16px; }
    .sep { border-top: 1px solid rgba(255,255,255,0.06); margin: 16px 0; }
    .section-label { font-size: 11px; font-weight: 600; color: #6b7280; text-transform: uppercase; letter-spacing: 1px; margin-bottom: 8px; }

    h1, h2, h3 { color: #e2e8f0 !important; }
    .stSelectbox label, .stSlider label { color: #94a3b8 !important; font-weight: 500; font-size: 13px; }

    /* Tab styling */
    .stTabs [data-baseweb="tab-list"] { gap: 0; }
    .stTabs [data-baseweb="tab"] { font-size: 13px; font-weight: 500; padding: 10px 20px; }

    /* Footer */
    .footer { text-align: center; color: #4b5563; font-size: 11px; margin-top: 40px; padding: 20px 0; border-top: 1px solid rgba(255,255,255,0.04); }
</style>
""", unsafe_allow_html=True)

# --- DATA LOADING ---
def load_data(file_pronostico=None, file_inventario=None, file_econometrico=None):
    if file_pronostico is not None:
        xls_pron = pd.ExcelFile(file_pronostico)
    else:
        local_files = [f for f in os.listdir('.') if "Pronosticos" in f and f.endswith('.xlsx')]
        xls_pron = pd.ExcelFile(local_files[0]) if local_files else None

    if file_inventario is not None:
        df_inv = pd.read_csv(file_inventario)
    else:
        local_csv = [f for f in os.listdir('.') if "inventario" in f.lower() and f.endswith('.csv')]
        df_inv = pd.read_csv(local_csv[0]) if local_csv else None

    if file_econometrico is not None:
        xls_econ = file_econometrico
    else:
        local_econ = [f for f in os.listdir('.') if any(k in f for k in ["Reporte", "PRECIOS", "OFERTAS", "econom"]) and f.endswith('.xlsx')]
        xls_econ = local_econ[0] if local_econ else None

    if xls_econ is not None:
        df_econ_raw = pd.read_excel(xls_econ, sheet_name=0)
        row_idx = df_econ_raw[df_econ_raw.iloc[:, 0] == 'Categoría'].index[0]
        df_ped = df_econ_raw.iloc[row_idx+1:row_idx+6, 0:5].copy()
        df_ped.columns = ['Category', 'PED', 'PED_Abs', 'Tipo_Demanda', 'Decision']
        df_ped['PED'] = df_ped['PED'].astype(float)
    else:
        df_ped = pd.DataFrame({
            'Category': ['Electronics', 'Clothing', 'Groceries', 'Toys', 'Furniture'],
            'PED': [0.2830, 0.1345, 0.0855, 0.5820, 0.6037]
        })
    return xls_pron, df_inv, df_ped

def calcular_escenario(df_forecast, df_ped, df_inv, descuento_pct, margen_pct, semanas_cob, prices, category_filter="Todas"):
    df = df_forecast.copy()
    if category_filter != "Todas":
        df = df[df['Category'] == category_filter]
    df = df.merge(df_ped[['Category', 'PED']], on='Category', how='left')
    if df_inv is not None and 'Stock_Seguridad' in df_inv.columns:
        df = df.merge(df_inv[['Category', 'Stock_Seguridad', 'Punto_Reorden', 'Brecha_Inventario']], on='Category', how='left')
    df['Precio_Base'] = df['Category'].map(prices).fillna(50.0)
    df['Delta_Unidades'] = df['Demanda_Base'] * df['PED'] * (descuento_pct / 100.0)
    df['Demanda_Ajustada'] = df['Demanda_Base'] + df['Delta_Unidades']
    df['Precio_Ajustado'] = df['Precio_Base'] * (1 - descuento_pct / 100.0)
    df['Ingreso_Ajustado'] = df['Demanda_Ajustada'] * df['Precio_Ajustado']
    df['Margen_Ajustado'] = df['Ingreso_Ajustado'] * margen_pct
    num_semanas = df['Semana_Futura'].nunique() if 'Semana_Futura' in df.columns else 26
    df['Demanda_Semanal_Promedio'] = df['Demanda_Ajustada'] / num_semanas
    df['Inventario_Requerido_Unidades'] = df['Demanda_Semanal_Promedio'] * semanas_cob
    df['Inventario_Requerido'] = df['Inventario_Requerido_Unidades'] * df['Precio_Ajustado']
    total_margen = df['Margen_Ajustado'].sum()
    df_unique_inv = df.groupby('Category')['Inventario_Requerido'].mean().reset_index()
    total_inventario = df_unique_inv['Inventario_Requerido'].sum()
    total_demanda = df['Demanda_Ajustada'].sum()
    total_delta = df['Delta_Unidades'].sum()
    ped_prom = df['PED'].mean()
    return df, {'margen': total_margen, 'inventario': total_inventario, 'demanda': total_demanda, 'delta': total_delta, 'ped': ped_prom}

def fmt(val):
    if abs(val) >= 1e6:
        return f"{val/1e6:.2f}M"
    elif abs(val) >= 1e3:
        return f"{val/1e3:.1f}K"
    else:
        return f"{val:.2f}"

CHART_LAYOUT = dict(
    paper_bgcolor='rgba(0,0,0,0)',
    plot_bgcolor='rgba(0,0,0,0)',
    font=dict(color="#94a3b8", size=12, family="Inter"),
    margin=dict(l=40, r=20, t=40, b=40),
    yaxis=dict(gridcolor='rgba(255,255,255,0.04)', zerolinecolor='rgba(255,255,255,0.06)'),
    xaxis=dict(gridcolor='rgba(255,255,255,0.04)'),
    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1, font=dict(size=11)),
)

COLORS = {
    'primary': '#60a5fa',
    'secondary': '#a78bfa',
    'success': '#34d399',
    'danger': '#f87171',
    'warning': '#fb923c',
    'muted': '#6b7280',
    'bar1': '#3b82f6',
    'bar2': '#8b5cf6',
}

# --- SIDEBAR ---
st.sidebar.markdown('<div class="sidebar-title">🛒 Retail Market</div>', unsafe_allow_html=True)
st.sidebar.markdown('<div class="sidebar-sub">Dashboard Prescriptivo</div>', unsafe_allow_html=True)
st.sidebar.markdown('<div class="sep"></div>', unsafe_allow_html=True)

st.sidebar.markdown('<div class="section-label">📁 Datos</div>', unsafe_allow_html=True)
up_pron = st.sidebar.file_uploader("Pronósticos (.xlsx)", type=["xlsx"])
up_inv = st.sidebar.file_uploader("Inventarios (.csv)", type=["csv"])
up_econ = st.sidebar.file_uploader("Econométrico (.xlsx)", type=["xlsx"])

xls_pron, df_inv, df_ped = load_data(up_pron, up_inv, up_econ)
if xls_pron is None or df_inv is None or df_ped is None:
    st.error("⚠️ Sube los archivos de datos en el panel lateral.")
    st.stop()

st.sidebar.markdown('<div class="sep"></div>', unsafe_allow_html=True)
st.sidebar.markdown('<div class="section-label">⚙️ Modelo</div>', unsafe_allow_html=True)
available_sheets = [s for s in xls_pron.sheet_names if s != 'Resumen_MAPE']
selected_sheet = st.sidebar.selectbox("Técnica de pronóstico", available_sheets, index=0)
df_pron_selected = pd.read_excel(xls_pron, sheet_name=selected_sheet)

forecast_cols = [c for c in df_pron_selected.columns if c.startswith('Demanda_')]
df_forecast = df_pron_selected.melt(
    id_vars=['Semana_Futura', 'Fecha'], value_vars=forecast_cols,
    var_name='Category', value_name='Demanda_Base'
)
df_forecast['Category'] = df_forecast['Category'].str.replace('Demanda_', '')
df_forecast['Fecha'] = pd.to_datetime(df_forecast['Fecha'])

st.sidebar.markdown('<div class="sep"></div>', unsafe_allow_html=True)
st.sidebar.markdown('<div class="section-label">🎛️ Sensibilidad</div>', unsafe_allow_html=True)
categories_list = ["Todas"] + list(df_forecast['Category'].unique())
selected_category = st.sidebar.selectbox("Categoría", categories_list)
descuento = st.sidebar.slider("Descuento (%)", 0, 30, 5, 1)
margen_pct = st.sidebar.slider("Margen (%)", 0.10, 0.90, 0.50, 0.05)
semanas_cobertura = st.sidebar.slider("Semanas cobertura", 1, 12, 4, 1)

prices = {'Groceries': 35.0, 'Clothing': 65.0, 'Toys': 45.0, 'Furniture': 120.0, 'Electronics': 150.0}

df_calc, kpis = calcular_escenario(df_forecast, df_ped, df_inv, descuento, margen_pct, semanas_cobertura, prices, selected_category)
_, kpi_con = calcular_escenario(df_forecast, df_ped, df_inv, 0, margen_pct, semanas_cobertura, prices, selected_category)
_, kpi_mod = calcular_escenario(df_forecast, df_ped, df_inv, 5, margen_pct, semanas_cobertura, prices, selected_category)
_, kpi_agr = calcular_escenario(df_forecast, df_ped, df_inv, 10, margen_pct, semanas_cobertura, prices, selected_category)

# --- HEADER ---
ds = "+" if kpis['delta'] >= 0 else ""
dc = "c-green" if kpis['delta'] >= 0 else "c-red"

st.markdown(f"""
<div style="margin-bottom:8px;">
    <span style="font-size:28px; font-weight:700; color:#e2e8f0;">📊 Sensibilidad Conjunta</span>
    <span style="font-size:13px; color:#6b7280; margin-left:16px;">Descuento {descuento}% · Margen {int(margen_pct*100)}% · {semanas_cobertura} sem.</span>
</div>
<div class="kpi-row">
    <div class="kpi"><div class="kpi-label">Margen</div><div class="kpi-val c-orange">{fmt(kpis['margen'])}</div></div>
    <div class="kpi"><div class="kpi-label">Inventario</div><div class="kpi-val c-blue">{fmt(kpis['inventario'])}</div></div>
    <div class="kpi"><div class="kpi-label">Demanda</div><div class="kpi-val c-purple">{fmt(kpis['demanda'])}</div></div>
    <div class="kpi"><div class="kpi-label">Delta uds.</div><div class="kpi-val {dc}">{ds}{fmt(kpis['delta'])}</div></div>
    <div class="kpi"><div class="kpi-label">PED</div><div class="kpi-val c-orange">{kpis['ped']:.2f}</div></div>
</div>
""", unsafe_allow_html=True)

# --- TABS ---
tab1, tab2, tab3, tab4 = st.tabs(["📊 Sensibilidad", "🎯 Escenarios", "📈 Elasticidad", "💡 Recomendación"])

# ============ TAB 1 ============
with tab1:
    df_cat = df_calc.groupby('Category').agg({'Demanda_Base':'sum','Demanda_Ajustada':'sum','Delta_Unidades':'sum'}).reset_index()

    c1, c2 = st.columns(2)
    with c1:
        fig1 = go.Figure()
        fig1.add_trace(go.Bar(x=df_cat['Category'], y=df_cat['Demanda_Base'], name='Base', marker_color=COLORS['bar1'],
                              text=[fmt(x) for x in df_cat['Demanda_Base']], textposition='auto', textfont=dict(size=11)))
        fig1.add_trace(go.Bar(x=df_cat['Category'], y=df_cat['Demanda_Ajustada'], name='Ajustada', marker_color=COLORS['bar2'],
                              text=[fmt(x) for x in df_cat['Demanda_Ajustada']], textposition='auto', textfont=dict(size=11)))
        fig1.update_layout(**CHART_LAYOUT, barmode='group', title=dict(text="Demanda Base vs Ajustada", font=dict(size=14)))
        st.plotly_chart(fig1, use_container_width=True)

    with c2:
        colors_delta = [COLORS['danger'] if x < 0 else COLORS['success'] for x in df_cat['Delta_Unidades']]
        fig2 = go.Figure()
        fig2.add_trace(go.Bar(x=df_cat['Category'], y=df_cat['Delta_Unidades'], marker_color=colors_delta,
                              text=[f"{'+' if x>=0 else ''}{fmt(x)}" for x in df_cat['Delta_Unidades']],
                              textposition='outside', textfont=dict(size=12, color='#e2e8f0')))
        fig2.update_layout(**CHART_LAYOUT, title=dict(text="Delta Unidades", font=dict(size=14)))
        st.plotly_chart(fig2, use_container_width=True)

    df_time = df_calc.groupby(['Fecha', 'Category'])['Demanda_Ajustada'].sum().reset_index()
    fig_t = px.line(df_time, x='Fecha', y='Demanda_Ajustada', color='Category', markers=True,
                    color_discrete_sequence=['#34d399','#fb923c','#60a5fa','#e879f9','#a3e635'])
    fig_t.update_layout(**CHART_LAYOUT, title=dict(text="Evolución Temporal — Demanda Ajustada", font=dict(size=14)), height=350)
    st.plotly_chart(fig_t, use_container_width=True)

# ============ TAB 2 ============
with tab2:
    st.markdown(f"""
    <div class="sc-grid">
        <div class="sc">
            <div class="sc-head">
                <span class="sc-name">Conservador</span>
                <span class="badge badge-blue">0%</span>
            </div>
            <div class="sc-row"><span class="sc-metric">Margen</span><span class="sc-val">{fmt(kpi_con['margen'])}</span></div>
            <div class="sc-row"><span class="sc-metric">Demanda</span><span class="sc-val">{fmt(kpi_con['demanda'])}</span></div>
            <div class="sc-row"><span class="sc-metric">Delta</span><span class="sc-val" style="color:#6b7280;">0</span></div>
            <div class="sc-row"><span class="sc-metric">Inventario</span><span class="sc-val">{fmt(kpi_con['inventario'])}</span></div>
        </div>
        <div class="sc sc-optimal">
            <div class="sc-head">
                <span class="sc-name">Moderado Selectivo</span>
                <span class="badge badge-green">5% ✓</span>
            </div>
            <div class="sc-row"><span class="sc-metric">Margen</span><span class="sc-val" style="color:#34d399;">{fmt(kpi_mod['margen'])}</span></div>
            <div class="sc-row"><span class="sc-metric">Demanda</span><span class="sc-val" style="color:#34d399;">{fmt(kpi_mod['demanda'])}</span></div>
            <div class="sc-row"><span class="sc-metric">Delta</span><span class="sc-val" style="color:#34d399;">+{fmt(kpi_mod['delta'])}</span></div>
            <div class="sc-row"><span class="sc-metric">Inventario</span><span class="sc-val">{fmt(kpi_mod['inventario'])}</span></div>
        </div>
        <div class="sc">
            <div class="sc-head">
                <span class="sc-name">Agresivo</span>
                <span class="badge badge-red">10%</span>
            </div>
            <div class="sc-row"><span class="sc-metric">Margen</span><span class="sc-val" style="color:#f87171;">{fmt(kpi_agr['margen'])}</span></div>
            <div class="sc-row"><span class="sc-metric">Demanda</span><span class="sc-val">{fmt(kpi_agr['demanda'])}</span></div>
            <div class="sc-row"><span class="sc-metric">Delta</span><span class="sc-val">+{fmt(kpi_agr['delta'])}</span></div>
            <div class="sc-row"><span class="sc-metric">Inventario</span><span class="sc-val" style="color:#f87171;">{fmt(kpi_agr['inventario'])}</span></div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    ce1, ce2 = st.columns(2)
    esc_df = pd.DataFrame({
        'Escenario': ['0%', '5%', '10%'],
        'Margen': [kpi_con['margen'], kpi_mod['margen'], kpi_agr['margen']],
        'Delta': [kpi_con['delta'], kpi_mod['delta'], kpi_agr['delta']]
    })
    esc_colors = [COLORS['primary'], COLORS['success'], COLORS['danger']]
    with ce1:
        fig_m = go.Figure(go.Bar(x=esc_df['Escenario'], y=esc_df['Margen'], marker_color=esc_colors,
                                 text=[fmt(x) for x in esc_df['Margen']], textposition='outside', textfont=dict(size=13, color='#e2e8f0')))
        fig_m.update_layout(**CHART_LAYOUT, title=dict(text="Margen por Escenario", font=dict(size=14)))
        st.plotly_chart(fig_m, use_container_width=True)
    with ce2:
        fig_d = go.Figure(go.Bar(x=esc_df['Escenario'], y=esc_df['Delta'], marker_color=esc_colors,
                                 text=[fmt(x) for x in esc_df['Delta']], textposition='outside', textfont=dict(size=13, color='#e2e8f0')))
        fig_d.update_layout(**CHART_LAYOUT, title=dict(text="Delta Unidades por Escenario", font=dict(size=14)))
        st.plotly_chart(fig_d, use_container_width=True)

# ============ TAB 3 ============
with tab3:
    df_ped_s = df_ped.sort_values('PED', ascending=True)
    ped_colors = [COLORS['success'] if p > 0.5 else COLORS['primary'] if p > 0.2 else COLORS['muted'] for p in df_ped_s['PED']]

    fig_p = go.Figure(go.Bar(
        y=df_ped_s['Category'], x=df_ped_s['PED'], orientation='h', marker_color=ped_colors,
        text=[f"{p:.4f}" for p in df_ped_s['PED']], textposition='outside', textfont=dict(size=13, color='#e2e8f0')
    ))
    fig_p.add_vline(x=1.0, line_dash="dash", line_color=COLORS['danger'],
                    annotation_text="|PED|=1", annotation_position="top right", annotation_font_color=COLORS['danger'])
    fig_p.update_layout(**CHART_LAYOUT, title=dict(text="Elasticidad Precio-Demanda (PED)", font=dict(size=14)),
                        xaxis=dict(title="PED", range=[0, 1.2], gridcolor='rgba(255,255,255,0.04)'), yaxis=dict(title=""), height=380)
    st.plotly_chart(fig_p, use_container_width=True)

    df_dec = df_ped[['Category', 'PED']].copy()
    df_dec['|PED|'] = df_dec['PED'].abs()
    df_dec['Sensibilidad'] = df_dec['|PED|'].apply(lambda x: '🟢 Alta' if x > 0.5 else '🔵 Baja' if x > 0.2 else '⚪ Rígida')
    df_dec['Descuento'] = df_dec['|PED|'].apply(lambda x: '5% selectivo' if x > 0.5 else 'No aplicar')
    st.dataframe(df_dec[['Category', '|PED|', 'Sensibilidad', 'Descuento']].style.format({'|PED|': '{:.4f}'}),
                 use_container_width=True, hide_index=True)

# ============ TAB 4 ============
with tab4:
    dem_pct = ((kpi_mod['demanda'] / kpi_con['demanda']) - 1) * 100
    mar_diff = kpi_mod['margen'] - kpi_con['margen']

    st.markdown(f"""
    <div class="rec-card">
        <div class="rec-title">✅ Descuento Selectivo del 5%</div>
        <div class="rec-sub">
            Aplicar solo a <span class="rec-highlight">Furniture (PED: 0.60)</span> y
            <span class="rec-highlight">Toys (PED: 0.58)</span>.<br><br>
            Incremento de demanda: <span class="rec-highlight">+{dem_pct:.1f}%</span> ·
            Variación de margen: <span class="rec-highlight">{fmt(mar_diff)}</span>
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.write("")
    resumen = pd.DataFrame({
        'Escenario': ['Conservador (0%)', 'Moderado (5%)', 'Agresivo (10%)'],
        'Descuento': ['0%', '5% Furniture+Toys', '10% General'],
        'Margen': [fmt(kpi_con['margen']), fmt(kpi_mod['margen']), fmt(kpi_agr['margen'])],
        'Demanda': [fmt(kpi_con['demanda']), fmt(kpi_mod['demanda']), fmt(kpi_agr['demanda'])],
        'Delta': ['0', f"+{fmt(kpi_mod['delta'])}", f"+{fmt(kpi_agr['delta'])}"],
        'Veredicto': ['Sin crecimiento', '✅ Óptimo', '⚠️ Riesgo']
    })
    st.dataframe(resumen, use_container_width=True, hide_index=True)

    st.markdown("""
    <div class="footer">
        EC3 · Modelamiento Predictivo de Datos (MPD) · Universidad de Lima · Sección 805<br>
        Random Forest TS · SARIMA · Holt-Winters · PCA · Regresión Log-Log · Inventario Probabilístico
    </div>
    """, unsafe_allow_html=True)
