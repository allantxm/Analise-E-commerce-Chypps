"""Chypps Tecnologia | Dashboard analítico (Streamlit)"""

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
TEXT, MUTED = "#EDF0F8", "#9AA3B8"
PALETA = [PURPLE, TEAL, AMBER, ROSE, BLUE, LIME]
ESCALA_ROXA = [[0, "#3A3278"], [1, PURPLE]]

@st.cache_data(show_spinner="Lendo e preparando a base...")
def carregar_base(conteudo: bytes, extensao: str) -> pd.DataFrame:
    return dados.carregar(conteudo, extensao=extensao)

st.sidebar.markdown("### 📊 Chypps Tecnologia")

caminho_xlsx = Path(__file__).resolve().parent / "fVendas.xlsx"
caminho_csv = Path(__file__).resolve().parent / "fVendas.csv"

upload = st.sidebar.file_uploader("Carregar planilha (Excel ou CSV)", type=["xlsx", "xls", "csv"])

try:
    if upload is not None:
        ext = upload.name.split(".")[-1]
        df = carregar_base(upload.getvalue(), extensao=ext)
        origem = upload.name
    elif caminho_xlsx.exists():
        df = carregar_base(caminho_xlsx.read_bytes(), extensao="xlsx")
        origem = caminho_xlsx.name
    elif caminho_csv.exists():
        df = carregar_base(caminho_csv.read_bytes(), extensao="csv")
        origem = caminho_csv.name
    else:
        st.warning("Coloque o arquivo fVendas.xlsx na mesma pasta do dashboard.")
        st.stop()
except Exception as erro:
    st.error(f"Erro ao carregar base: {erro}")
    st.stop()

if df.empty:
    st.warning("A base está vazia.")
    st.stop()

st.sidebar.divider()
st.sidebar.subheader("Filtros")
data_min, data_max = df["Data"].min().date(), df["Data"].max().date()
selecao = st.sidebar.date_input("Período", value=(data_min, data_max), min_value=data_min, max_value=data_max)

inicio, fim = (selecao[0], selecao[1]) if isinstance(selecao, (tuple, list)) and len(selecao) == 2 else (data_min, data_max)

lojas = sorted(df["Loja virtual"].dropna().unique().tolist())
canais = st.sidebar.multiselect("Conta / Loja virtual", lojas, default=lojas)

ufs = sorted(df["Estado"].dropna().unique().tolist())
estados = st.sidebar.multiselect("Estado (UF)", ufs, default=ufs)

f = dados.filtrar(df, inicio, fim, canais=canais, estados=estados)
ind = dados.indicadores(f)

st.title("📊 Painel de Vendas - Vali Transfers & Chypps")
st.markdown("---")

col1, col2, col3, col4 = st.columns(4)
col1.metric("Faturamento Líquido", brl(ind["faturamento"]))
col2.metric("Faturamento Bruto", brl(ind["faturamento_bruto"]))
col3.metric("Total de Pedidos", inteiro(ind["pedidos"]))
col4.metric("Ticket Médio", brl(ind["ticket_medio"]))

st.markdown("---")
st.subheader("Faturamento por Conta / Loja Virtual")
canal = f.groupby("Loja virtual", as_index=False).agg(Faturamento=("Preço Total", "sum"))
fig = px.pie(canal, names="Loja virtual", values="Faturamento", hole=0.6, color_discrete_sequence=PALETA)
st.plotly_chart(fig, use_container_width=True)