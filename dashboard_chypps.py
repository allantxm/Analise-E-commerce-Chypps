"""
Chypps Tecnologia | Dashboard analítico (Streamlit)
Integrado com as medidas de Faturamento Líquido, Bruto, Mês Atual, Mês Anterior e % MoM.
"""

import html
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

import dados
from dados import brl, inteiro, pct


st.set_page_config(
    page_title="Chypps Tecnologia | Inteligência Comercial",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

BG, PANEL, PANEL_HI, LINE_SOLID = "#0A0E19", "#121929", "#1A2338", "#222C42"
LINE, GRID = "rgba(255,255,255,0.07)", "rgba(255,255,255,0.06)"
PURPLE, TEAL, AMBER, ROSE = "#8F7DFF", "#38D6C4", "#F4B942", "#FF6B8B"
BLUE, LIME = "#5AA2FF", "#9BDB7A"
TEXT, MUTED, CINZA = "#EDF0F8", "#9AA3B8", "#566078"
PALETA = [PURPLE, TEAL, AMBER, ROSE, BLUE, LIME]
ESCALA_ROXA = [[0, "#3A3278"], [1, PURPLE]]
FONTE = "Manrope, 'Segoe UI', system-ui, -apple-system, sans-serif"

PONTILHADO = (
    "url(\"data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='26' height='26'%3E"
    "%3Ccircle cx='1.5' cy='1.5' r='1' fill='white' fill-opacity='0.07'/%3E%3C/svg%3E\")"
)

CSS = """
@import url('https://fonts.googleapis.com/css2?family=Manrope:wght@400;500;600;700;800&display=swap');
:root { color-scheme: dark; }
html, body, .stApp, .stApp :is(p, h1, h2, h3, h4, label, li, button, input, textarea, td, th),
.stApp [data-testid="stMetricValue"], .stApp [class*="vt-"] {
    font-family: 'Manrope', 'Segoe UI', system-ui, -apple-system, sans-serif;
}
.stApp { font-variant-numeric: tabular-nums; isolation: isolate; color: __TEXT__; background-color: __BG__;
    background-image: radial-gradient(1100px 620px at 6% -8%, rgba(143,125,255,.14), transparent 62%),
    radial-gradient(900px 560px at 104% 2%, rgba(56,214,196,.08), transparent 58%),
    radial-gradient(900px 600px at 50% 120%, rgba(90,162,255,.06), transparent 60%);
    background-attachment: fixed; }
section[data-testid="stSidebar"] { background:rgba(11,16,28,.84); -webkit-backdrop-filter:blur(16px);
    backdrop-filter:blur(16px); border-right:1px solid __LINE__; }
.vt-hero { display:flex; flex-wrap:wrap; align-items:center; justify-content:space-between;
    gap:16px 28px; margin:0 0 22px; padding-bottom:20px; border-bottom:1px solid __LINE__; }
.vt-title { font-size:1.85rem; font-weight:800; letter-spacing:-.03em; color:__TEXT__; }
.vt-subtitle { margin-top:4px; font-size:.95rem; color:__MUTED__; }
.vt-kpis { display:grid; grid-template-columns:repeat(4, minmax(0, 1fr)); gap:14px; }
.vt-kpi { padding:18px 20px; border-radius:18px; background:linear-gradient(160deg,rgba(26,35,56,.92),rgba(18,25,41,.92)); border:1px solid __LINE__; }
.vt-kpi-label { font-size:.84rem; font-weight:600; color:__MUTED__; margin-bottom:6px; }
.vt-kpi-value { font-size:1.6rem; font-weight:800; color:__TEXT__; }
.vt-kpi-sub { font-size:.8rem; color:__MUTED__; margin-top:4px; }
"""

_TOKENS = {
    "__PONTILHADO__": PONTILHADO, "__BG__": BG, "__PANELHI__": PANEL_HI,
    "__PANEL__": PANEL, "__LINESOLID__": LINE_SOLID, "__LINE__": LINE,
    "__PURPLE__": PURPLE, "__TEAL__": TEAL, "__AMBER__": AMBER,
    "__ROSE__": ROSE, "__TEXT__": TEXT, "__MUTED__": MUTED,
}
_css = CSS
for _chave, _valor in _TOKENS.items():
    _css = _css.replace(_chave, _valor)
st.markdown(f"<style>{_css}</style>", unsafe_allow_html=True)


def _versao():
    try:
        return tuple(int(p) for p in st.__version__.split(".")[:2])
    except (ValueError, AttributeError):
        return (0, 0)


_LARGURA = {"width": "stretch"} if _versao() >= (1, 50) else {"use_container_width": True}


def _h(texto) -> str:
    return html.escape(str(texto)).replace("$", "&#36;")


def grafico(fig, altura=360):
    fig.update_layout(
        height=altura, template="plotly_dark",
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color=TEXT, family=FONTE, size=12),
        margin=dict(l=8, r=8, t=24, b=8),
    )
    st.plotly_chart(fig, **_LARGURA)


def tabela(df, formatos=None, altura=None):
    visual = df.style.format(formatos, na_rep="—") if formatos else df
    extra = {"height": altura} if altura else {}
    st.dataframe(visual, hide_index=True, **_LARGURA, **extra)


@st.cache_data(show_spinner="Lendo base...")
def carregar_base(conteudo: bytes) -> pd.DataFrame:
    return dados.carregar(conteudo, extensao="xlsx")


st.sidebar.markdown("### 📊 Chypps Tecnologia")
caminho_xlsx = Path(__file__).resolve().parent / "fVendas.xlsx"

upload = st.sidebar.file_uploader("Carregar fVendas.xlsx", type=["xlsx"])

try:
    if upload is not None:
        df = carregar_base(upload.getvalue())
        origem = upload.name
    elif caminho_xlsx.exists():
        df = carregar_base(caminho_xlsx.read_bytes())
        origem = caminho_xlsx.name
    else:
        st.warning("Arquivo fVendas.xlsx não encontrado.")
        st.stop()
except Exception as erro:
    st.error(f"Erro ao carregar base: {erro}")
    st.stop()

st.sidebar.divider()
st.sidebar.subheader("Filtros")
datas_validas = df["Data"].dropna()
data_min, data_max = datas_validas.min().date(), datas_validas.max().date()

selecao = st.sidebar.date_input(
    "Período", 
    value=(data_min, data_max), 
    min_value=data_min, 
    max_value=data_max
)

if isinstance(selecao, tuple):
    if len(selecao) == 2:
        inicio, fim = selecao
    elif len(selecao) == 1:
        inicio = fim = selecao[0]
    else:
        inicio, fim = data_min, data_max
else:
    inicio = fim = selecao

lojas = sorted(df["Loja virtual"].dropna().unique().tolist())
canais = st.sidebar.multiselect("Canal de venda (Conta)", lojas, default=lojas)

ufs = sorted(df["Estado"].dropna().unique().tolist())
estados = st.sidebar.multiselect("Estado", ufs, default=ufs)

f = dados.filtrar(df, inicio, fim, None, canais, estados, None)

fat_liq_atual, fat_liq_ant, fat_bruto_atual, fat_bruto_ant, variacao_mom = dados.calcular_metricas_mes_atual_anterior(df, canais, estados)

if variacao_mom > 0:
    seta_mom = f"📈 +{variacao_mom * 100:.2f}%".replace(".", ",")
elif variacao_mom < 0:
    seta_mom = f"📉 {variacao_mom * 100:.2f}%".replace(".", ",")
else:
    seta_mom = "⚪ 0,00%"

st.markdown(
    f"<div class='vt-hero'><div><div class='vt-title'>Inteligência Comercial - Vali Transfers</div>"
    f"<div class='vt-subtitle'>Acompanhamento de Faturamento e Vendas (Base: {_h(origem)})</div></div></div>",
    unsafe_allow_html=True,
)

st.markdown("### 📌 Indicadores do Mês Atual vs Mês Anterior (DAX Equivalente)")
col1, col2, col3, col4 = st.columns(4)

with col1:
    st.markdown(f"""
        <div class='vt-kpi'>
            <div class='vt-kpi-label'>Faturamento Bruto (Atual)</div>
            <div class='vt-kpi-value'>{brl(fat_bruto_atual)}</div>
            <div class='vt-kpi-sub'>Anterior: {brl(fat_bruto_ant)}</div>
        </div>
    """, unsafe_allow_html=True)

with col2:
    st.markdown(f"""
        <div class='vt-kpi'>
            <div class='vt-kpi-label'>Faturamento Líquido (Atual)</div>
            <div class='vt-kpi-value'>{brl(fat_liq_atual)}</div>
            <div class='vt-kpi-sub'>Anterior: {brl(fat_liq_ant)}</div>
        </div>
    """, unsafe_allow_html=True)

with col3:
    st.markdown(f"""
        <div class='vt-kpi'>
            <div class='vt-kpi-label'>Variação % MoM (Líquido)</div>
            <div class='vt-kpi-value'>{seta_mom}</div>
            <div class='vt-kpi-sub'>Comparativo mensal fixo</div>
        </div>
    """, unsafe_allow_html=True)

with col4:
    ind = dados.indicadores(f)
    st.markdown(f"""
        <div class='vt-kpi'>
            <div class='vt-kpi-label'>Ticket Médio (Filtro)</div>
            <div class='vt-kpi-value'>{brl(ind['ticket_medio'])}</div>
            <div class='vt-kpi-sub'>Total Pedidos: {inteiro(ind['pedidos'])}</div>
        </div>
    """, unsafe_allow_html=True)

st.markdown("---")

aba1, aba2, aba3 = st.tabs(["Visão Geral", "Produtos (Curva ABC)", "Geografia"])

with aba1:
    diaria = dados.serie_diaria(f)
    fig = px.bar(diaria, x="Data", y="Faturamento", title="Evolução Diária do Faturamento Líquido", color_discrete_sequence=[PURPLE])
    grafico(fig, 360)

with aba2:
    abc = dados.curva_abc(f)
    st.subheader("Curva ABC de Produtos")
    tabela(abc, {"Faturamento": brl, "% do total": "{:.2f}%", "% acumulado": "{:.2f}%"})

with aba3:
    geo = f.groupby("Estado", as_index=False).agg(Faturamento=("Faturamento Liquido", "sum"), Pedidos=("N° do Pedido", "nunique"))
    st.subheader("Vendas por Estado")
    tabela(geo.sort_values("Faturamento", ascending=False), {"Faturamento": brl, "Pedidos": inteiro})