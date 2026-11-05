"""
Chypps Tecnologia | Dashboard analítico (Streamlit)

A base Excel/CSV é consumida SOMENTE PARA LEITURA.
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

# ---------- Identidade visual ----------
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
.stApp::before { content:""; position:fixed; inset:0; z-index:-1; pointer-events:none;
    background-image: __PONTILHADO__;
    -webkit-mask-image:linear-gradient(180deg,rgba(0,0,0,.95) 0%,rgba(0,0,0,0) 72%);
    mask-image:linear-gradient(180deg,rgba(0,0,0,.95) 0%,rgba(0,0,0,0) 72%); }
[data-testid="stAppViewContainer"], [data-testid="stMain"], .main { background:transparent; }
header[data-testid="stHeader"] { background:transparent; }
footer, [data-testid="stFooter"] { visibility:hidden; }
.block-container { max-width:1500px; padding-top:2.2rem; padding-bottom:3rem; }
section[data-testid="stSidebar"] { background:rgba(11,16,28,.84); -webkit-backdrop-filter:blur(16px);
    backdrop-filter:blur(16px); border-right:1px solid __LINE__; }
section[data-testid="stSidebar"] > div { background:transparent; }
section[data-testid="stSidebar"] hr { border-color:__LINE__; }
.vt-brand { display:flex; align-items:center; gap:12px; margin:2px 0 16px; }
.vt-brand .vt-mark { width:38px; height:38px; border-radius:12px; font-size:1.1rem; }
.vt-brand b { display:block; font-size:1.05rem; letter-spacing:-.01em; color:__TEXT__; }
.vt-brand span { color:__MUTED__; font-size:.8rem; }
.vt-side-count { margin-top:8px; padding:10px 12px; border-radius:12px; font-size:.8rem;
    color:__MUTED__; background:rgba(255,255,255,.04); border:1px solid __LINE__; }
.vt-side-count b { color:__TEXT__; }
.vt-hero { display:flex; flex-wrap:wrap; align-items:center; justify-content:space-between;
    gap:16px 28px; margin:0 0 22px; padding-bottom:20px; border-bottom:1px solid __LINE__; }
.vt-hero-id { display:flex; align-items:center; gap:16px; }
.vt-mark { width:46px; height:46px; flex:none; display:grid; place-items:center; border-radius:14px;
    font-weight:800; font-size:1.35rem; color:__BG__; background:linear-gradient(135deg,__PURPLE__,__TEAL__);
    box-shadow:0 8px 24px rgba(143,125,255,.28); }
.vt-title { font-size:1.85rem; font-weight:800; letter-spacing:-.03em; line-height:1.1; color:__TEXT__; }
.vt-subtitle { margin-top:4px; font-size:.95rem; color:__MUTED__; }
.vt-chips { display:flex; flex-wrap:wrap; gap:8px; }
.vt-chip { padding:6px 12px; border-radius:999px; font-size:.8rem; font-weight:600; color:__TEXT__;
    background:rgba(255,255,255,.05); border:1px solid __LINE__; }
.vt-chip-soft { color:__MUTED__; font-weight:500; }
.vt-kpis { display:grid; grid-template-columns:1.7fr 1fr 1fr 1fr; gap:14px; }
.vt-kpi { position:relative; display:flex; flex-direction:column; gap:6px; overflow:hidden;
    min-height:128px; padding:18px 20px 16px; border-radius:18px;
    background:linear-gradient(160deg,rgba(26,35,56,.92),rgba(18,25,41,.92)); border:1px solid __LINE__; }
.vt-kpi-hero { min-height:150px; }
.vt-kpi-label { display:flex; align-items:center; gap:8px; font-size:.84rem; font-weight:600; color:__MUTED__; }
.vt-kpi-label i { flex:none; width:8px; height:8px; border-radius:50%; background:var(--cor); }
.vt-kpi-value { font-size:clamp(1.3rem,1.75vw,1.7rem); font-weight:800; letter-spacing:-.02em;
    line-height:1.15; color:__TEXT__; overflow-wrap:anywhere; }
.vt-kpi-hero .vt-kpi-value { font-size:clamp(1.8rem,2.6vw,2.5rem); }
.vt-kpi-sub { font-size:.8rem; color:__MUTED__; }
.vt-chip-ok, .vt-chip-bad { display:inline-block; padding:2px 9px; border-radius:999px;
    font-size:.78rem; font-weight:700; }
.vt-chip-ok { background:rgba(56,214,196,.14); color:__TEAL__; }
.vt-chip-bad { background:rgba(255,107,139,.14); color:__ROSE__; }
.vt-spark { display:block; width:100%; height:56px; margin-top:auto; }
.has-help { text-decoration:underline dotted rgba(255,255,255,.35); text-underline-offset:3px; cursor:help; }
.vt-strip { display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:1px; overflow:hidden;
    margin:14px 0 8px; border-radius:16px; background:__LINESOLID__; border:1px solid __LINESOLID__; }
.vt-cell { padding:14px 18px; background:__PANEL__; }
.vt-cell-label { font-size:.8rem; font-weight:600; color:__MUTED__; }
.vt-cell-value { margin-top:4px; font-size:1.2rem; font-weight:700; letter-spacing:-.01em; color:__TEXT__; }
.vt-cell-value.warn { color:__AMBER__; }
.vt-note { margin:8px 2px 0; font-size:.8rem; line-height:1.5; color:__MUTED__; }
.vt-insights { margin:18px 0 8px; padding:16px 20px; border-radius:16px;
    background:rgba(18,25,41,.85); border:1px solid __LINE__; }
.vt-insights-title { margin-bottom:6px; font-size:.95rem; font-weight:700; color:__TEXT__; }
.vt-insights ul { margin:0; padding:0; list-style:none; columns:2 340px; column-gap:36px; }
.vt-insights li { position:relative; padding:5px 0 5px 18px; break-inside:avoid;
    font-size:.92rem; line-height:1.5; color:__TEXT__; }
.vt-insights li::before { content:""; position:absolute; left:0; top:.95em; width:7px; height:7px;
    border-radius:50%; background:__PURPLE__; }
div[data-baseweb="tab-list"] { gap:4px; }
button[data-baseweb="tab"] { height:46px; padding:0 16px; color:__MUTED__; }
button[data-baseweb="tab"] p { font-weight:600; font-size:.95rem; }
button[data-baseweb="tab"][aria-selected="true"] { color:__TEXT__; }
div[data-baseweb="tab-highlight"] { height:3px; border-radius:3px; background:__PURPLE__; }
div[data-baseweb="tab-border"] { background:__LINESOLID__; }
.stApp h3 { font-size:1.08rem; font-weight:700; letter-spacing:-.01em; }
[data-testid="stCaptionContainer"] { color:__MUTED__; }
[data-testid="stPlotlyChart"] { box-sizing:border-box; padding:10px 12px 4px; border-radius:16px;
    background:rgba(18,25,41,.78); border:1px solid __LINE__; }
[data-testid="stDataFrame"] { overflow:hidden; border-radius:14px; border:1px solid __LINE__; }
[data-testid="stMetric"] { padding:14px 18px; border-radius:14px; background:__PANEL__; border:1px solid __LINE__; }
[data-testid="stMetricLabel"] { color:__MUTED__; }
[data-testid="stMetricLabel"] p { font-size:.84rem; font-weight:600; }
[data-testid="stMetricValue"] { color:__TEXT__; font-size:1.5rem; font-weight:800; letter-spacing:-.02em; }
[data-testid="stAlert"] { border-radius:14px; }
div[data-baseweb="select"] > div, div[data-baseweb="input"], div[data-baseweb="base-input"] {
    border-radius:10px; background-color:#0E1424; border-color:__LINE__; }
span[data-baseweb="tag"] { border-radius:8px; color:__TEXT__; background-color:rgba(143,125,255,.18);
    border:1px solid rgba(143,125,255,.35); }
[data-testid="stFileUploaderDropzone"] { border-radius:12px; background:#0E1424;
    border:1px dashed rgba(255,255,255,.16); }
[data-testid="stDownloadButton"] button { padding:8px 18px; border-radius:10px; font-weight:700; color:__TEXT__;
    background:rgba(143,125,255,.16); border:1px solid rgba(143,125,255,.5); }
[data-testid="stDownloadButton"] button:hover { background:rgba(143,125,255,.3); border-color:__PURPLE__; }
.vt-footer { display:flex; flex-wrap:wrap; justify-content:space-between; gap:6px 24px;
    margin-top:36px; padding-top:16px; border-top:1px solid __LINE__; font-size:.8rem; color:__MUTED__; }
@media (max-width:1000px) { .vt-kpis { grid-template-columns:repeat(2,minmax(0,1fr)); }
    .vt-kpi-hero { grid-column:span 2; } .vt-strip { grid-template-columns:repeat(2,minmax(0,1fr)); } }
@media (max-width:560px) { .vt-kpis { grid-template-columns:minmax(0,1fr); }
    .vt-kpi-hero { grid-column:span 1; } .vt-title { font-size:1.5rem; } }
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


# ---------- Compatibilidade entre versões do Streamlit ----------
def _versao():
    try:
        return tuple(int(p) for p in st.__version__.split(".")[:2])
    except (ValueError, AttributeError):
        return (0, 0)


_LARGURA = {"width": "stretch"} if _versao() >= (1, 50) else {"use_container_width": True}


# ---------- Componentes visuais ----------
def _h(texto) -> str:
    return html.escape(str(texto)).replace("$", "&#36;")


def grafico(fig, altura=360):
    tem_titulo = bool(fig.layout.title.text)
    fig.update_layout(
        height=altura, template="plotly_dark",
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color=TEXT, family=FONTE, size=12), separators=",.",
        margin=dict(l=8, r=8, t=58 if tem_titulo else 24, b=8),
        legend=dict(title_text="", bgcolor="rgba(0,0,0,0)", font=dict(color=MUTED, size=12)),
        hoverlabel=dict(bgcolor=PANEL_HI, bordercolor="rgba(255,255,255,0.18)",
                        font=dict(color=TEXT, family=FONTE)),
        uniformtext=dict(minsize=11, mode="hide"),
    )
    if tem_titulo:
        fig.update_layout(title=dict(font=dict(size=14, color=TEXT), x=0, xanchor="left", y=0.96))
    eixo = dict(gridcolor=GRID, zeroline=False, linecolor=LINE, tickfont=dict(color=MUTED),
                title_font=dict(color=MUTED, size=12), title_standoff=8)
    fig.update_xaxes(**eixo)
    fig.update_yaxes(**eixo)
    fig.update_traces(marker=dict(line=dict(color=BG, width=2)), selector=dict(type="pie"))
    try:
        fig.update_layout(barcornerradius=6)
    except ValueError:
        pass
    st.plotly_chart(fig, **_LARGURA)


def centro_rosca(fig, valor, rotulo):
    fig.add_annotation(
        xref="paper", yref="paper", x=0.5, y=0.5, showarrow=False, align="center",
        text=f"<b>{_h(valor)}</b><br><span style='font-size:11px;color:{MUTED}'>{_h(rotulo)}</span>",
        font=dict(size=16, color=TEXT),
    )


def tabela(df, formatos=None, altura=None):
    visual = df.style.format(formatos, na_rep="—") if formatos else df
    extra = {"height": altura} if altura else {}
    st.dataframe(visual, hide_index=True, **_LARGURA, **extra)


def cabecalho(inicio, fim, origem):
    st.markdown(
        "<div class='vt-hero'><div class='vt-hero-id'><div class='vt-mark'>C</div><div>"
        "<div class='vt-title'>Chypps Tecnologia</div>"
        "<div class='vt-subtitle'>Inteligência comercial: vendas, canais, produtos e rentabilidade</div>"
        "</div></div><div class='vt-chips'>"
        f"<span class='vt-chip'>Período: {inicio:%d/%m/%Y} a {fim:%d/%m/%Y}</span>"
        f"<span class='vt-chip'>Fonte: {_h(origem)}</span>"
        "<span class='vt-chip vt-chip-soft'>Base de vendas consolidada</span>"
        "</div></div>",
        unsafe_allow_html=True,
    )


def sparkline(valores, cor):
    v = [float(x) for x in valores if pd.notna(x)]
    if len(v) < 2:
        return ""
    largura, altura = 240, 56
    menor, maior = min(v), max(v)
    faixa = (maior - menor) or 1.0
    passo = largura / (len(v) - 1)
    pontos = " ".join(
        f"{i * passo:.1f},{altura - 4 - (x - menor) / faixa * (altura - 8):.1f}"
        for i, x in enumerate(v)
    )
    return (
        f"<svg class='vt-spark' viewBox='0 0 {largura} {altura}' preserveAspectRatio='none' aria-hidden='true'>"
        f"<polygon points='0,{altura} {pontos} {largura},{altura}' fill='{cor}' fill-opacity='0.16'/>"
        f"<polyline points='{pontos}' fill='none' stroke='{cor}' stroke-width='2' stroke-linejoin='round' "
        "stroke-linecap='round' vector-effect='non-scaling-stroke'/></svg>"
    )


def card_kpi(rotulo, valor, sub="", ajuda="", cor=PURPLE, chip="", chip_classe="vt-chip-ok",
             extra="", destaque=False):
    classe = "vt-kpi vt-kpi-hero" if destaque else "vt-kpi"
    rot = (f"<span class='has-help' title='{_h(ajuda)}'>{_h(rotulo)}</span>"
           if ajuda else f"<span>{_h(rotulo)}</span>")
    partes = [
        f"<div class='{classe}' style='--cor:{cor}'>",
        f"<div class='vt-kpi-label'><i></i>{rot}</div>",
        f"<div class='vt-kpi-value'>{_h(valor)}</div>",
    ]
    if chip:
        partes.append(f"<div><span class='{chip_classe}'>{_h(chip)}</span></div>")
    if sub:
        partes.append(f"<div class='vt-kpi-sub'>{_h(sub)}</div>")
    partes += [extra, "</div>"]
    return "".join(partes)


def faixa_kpis(itens):
    celulas = []
    for rotulo, valor, ajuda, tom in itens:
        rot = (f"<span class='has-help' title='{_h(ajuda)}'>{_h(rotulo)}</span>"
               if ajuda else f"<span>{_h(rotulo)}</span>")
        celulas.append(
            f"<div class='vt-cell'><div class='vt-cell-label'>{rot}</div>"
            f"<div class='vt-cell-value {tom}'>{_h(valor)}</div></div>"
        )
    return "<div class='vt-strip'>" + "".join(celulas) + "</div>"


# ---------- Carga da base (somente leitura) ----------
@st.cache_data(show_spinner="Lendo e preparando a base...")
def carregar_base(conteudo: bytes, extensao: str) -> pd.DataFrame:
    return dados.carregar(conteudo, extensao=extensao)


st.sidebar.markdown(
    "<div class='vt-brand'><div class='vt-mark'>C</div><div><b>Chypps Tecnologia</b>"
    "<span>Painel de inteligência comercial</span></div></div>",
    unsafe_allow_html=True,
)

# Caminhos padrão suportados (.xlsx ou .csv)
caminho_xlsx = Path(__file__).resolve().parent / "fVendas.xlsx"
caminho_csv = Path(__file__).resolve().parent / "fVendas.csv"

upload = st.sidebar.file_uploader(
    "Carregar outra exportação (Excel ou CSV)", type=["xlsx", "xls", "csv"],
    help="A leitura é temporária, em memória. O arquivo original nunca é modificado.",
)

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
        st.warning(
            "Não encontrei o arquivo fVendas.xlsx ou fVendas.csv ao lado do dashboard. "
            "Coloque a planilha na mesma pasta ou envie uma exportação pelo campo da barra lateral."
        )
        st.stop()
except Exception as erro:
    st.error(f"Não foi possível carregar a base: {erro}")
    st.stop()

if df.empty:
    st.warning("A base não contém registros.")
    st.stop()


# ---------- Filtros ----------
st.sidebar.divider()
st.sidebar.subheader("Filtros")
datas_validas = df["Data"].dropna()
if datas_validas.empty:
    st.error("A coluna Data não contém datas válidas.")
    st.stop()

data_min, data_max = datas_validas.min().date(), datas_validas.max().date()
selecao = st.sidebar.date_input(
    "Período", value=(data_min, data_max), min_value=data_min, max_value=data_max,
    format="DD/MM/YYYY",
)
if isinstance(selecao, (tuple, list)):
    if len(selecao) == 2:
        inicio, fim = selecao
    elif len(selecao) == 1:
        inicio = fim = selecao[0]
    else:
        inicio, fim = data_min, data_max
else:
    inicio = fim = selecao

todas_situacoes = sorted(df["Situação"].dropna().unique().tolist())
padrao_situacoes = [s for s in todas_situacoes if s.casefold() != "cancelado"] or todas_situacoes
situacoes = st.sidebar.multiselect("Situação do pedido", todas_situacoes, default=padrao_situacoes)

lojas = sorted(df["Loja virtual"].dropna().unique().tolist())
canais = st.sidebar.multiselect("Canal de venda", lojas, default=lojas)

ufs = sorted(df["Estado"].dropna().unique().tolist())
estados = st.sidebar.multiselect("Estado", ufs, default=ufs)

produtos_lista = sorted(df["Descrição"].dropna().unique().tolist())
produtos = st.sidebar.multiselect(
    "Produto", produtos_lista, default=[], placeholder="Todos os produtos"
)
meta = st.sidebar.number_input(
    "Meta de faturamento do período (R$)", min_value=0.0, value=0.0, step=10000.0,
    help="Opcional. Se informada, aparece o acompanhamento da meta na Visão geral.",
)

f = dados.filtrar(df, inicio, fim, situacoes, canais, estados, produtos)
f_todas = dados.filtrar(df, inicio, fim, None, canais, estados, produtos)

st.sidebar.markdown(
    f"<div class='vt-side-count'><b>{_h(inteiro(len(f)))}</b> linhas e "
    f"<b>{_h(inteiro(f['N° do Pedido'].nunique()))}</b> pedidos após os filtros</div>",
    unsafe_allow_html=True,
)

if f.empty:
    cabecalho(inicio, fim, origem)
    st.info("Nenhum registro corresponde aos filtros selecionados. Ajuste o período ou os filtros.")
    st.stop()

ind = dados.indicadores(f)
taxa_canc = dados.taxa_cancelamento(f_todas)
diaria = dados.serie_diaria(f)


# ---------- Cabeçalho e indicadores ----------
cabecalho(inicio, fim, origem)
margem_ok = pd.notna(ind["margem_pct"])

st.markdown(
    "<div class='vt-kpis'>"
    + card_kpi(
        "Faturamento", brl(ind["faturamento"]), "Média móvel de 7 dias no período",
        cor=PURPLE, destaque=True,
        extra=sparkline(diaria["Média móvel 7 dias"], PURPLE) if not diaria.empty else "",
    )
    + card_kpi("Pedidos", inteiro(ind["pedidos"]), "Pedidos distintos no período", cor=BLUE)
    + card_kpi("Ticket médio", brl(ind["ticket_medio"]), "Faturamento ÷ pedidos", cor=AMBER,
               ajuda="Faturamento ÷ pedidos distintos.")
    + card_kpi(
        "Margem bruta est.", brl(ind["margem_bruta"]),
        "" if margem_ok else "Sem custo informado no recorte", cor=TEAL,
        chip=f"{pct(ind['margem_pct'])} de margem" if margem_ok else "",
        chip_classe="vt-chip-ok" if (not margem_ok or ind["margem_pct"] >= 0) else "vt-chip-bad",
        ajuda="Só considera linhas com custo informado.",
    )
    + "</div>",
    unsafe_allow_html=True,
)

st.markdown(
    faixa_kpis([
        ("Itens vendidos", inteiro(ind["itens"]), "", ""),
        ("Itens por pedido", f"{ind['itens_por_pedido']:.1f}".replace(".", ","), "", ""),
        ("Taxa de cancelamento", pct(taxa_canc),
         "Pedidos cancelados ÷ pedidos totais, nos demais filtros.", ""),
        ("Frete (por pedido)", brl(ind["frete_total"]),
         "Frete é do pedido: conta uma vez por pedido, não por item.", ""),
        ("Pedidos com frete zero", pct(ind["pct_frete_gratis"]), "", ""),
        ("Descontos (por pedido)", brl(ind["desconto_total"]), "", ""),
        ("NF-e pendentes", inteiro(ind["nfe_pendentes"]),
         "Pedidos com nota fiscal em situação 'Pendente'.",
         "warn" if ind["nfe_pendentes"] else ""),
        ("Cobertura de custo", pct(ind["cobertura_custo"], 0),
         "Parcela do faturamento com preço de custo informado.", ""),
    ]),
    unsafe_allow_html=True,
)

frases = dados.destaques(f, f_todas)
if frases:
    itens_html = "".join(f"<li>{_h(frase)}</li>" for frase in frases)
    st.markdown(
        f"<div class='vt-insights'><div class='vt-insights-title'>Destaques do período</div>"
        f"<ul>{itens_html}</ul></div>",
        unsafe_allow_html=True,
    )

aba_geral, aba_produtos, aba_canais, aba_geo, aba_pedidos, aba_dados = st.tabs(
    ["Visão geral", "Produtos", "Canais", "Geografia", "Pedidos e status", "Dados e exportação"]
)


# ======================= VISÃO GERAL =======================
with aba_geral:
    esq, dir_ = st.columns([1.3, 1])
    with esq:
        st.subheader("Evolução diária")
        fig = go.Figure()
        fig.add_trace(go.Bar(
            x=diaria["Data"], y=diaria["Faturamento"], name="Faturamento do dia",
            marker_color=PURPLE, opacity=0.5, customdata=diaria["Pedidos"],
            hovertemplate="%{x|%d/%m}: R$ %{y:,.2f}<br>%{customdata} pedidos<extra></extra>",
        ))
        fig.add_trace(go.Scatter(
            x=diaria["Data"], y=diaria["Média móvel 7 dias"], name="Média móvel de 7 dias",
            mode="lines", line=dict(color=TEAL, width=3),
        ))
        fig.update_yaxes(title="Faturamento", tickprefix="R$ ")
        fig.update_xaxes(tickformat="%d/%m")
        fig.update_layout(legend=dict(orientation="h", y=1.14, x=0), bargap=0.25)
        grafico(fig, 360)

    with dir_:
        st.subheader("Faturamento por canal")
        canal = f.groupby("Loja virtual", as_index=False).agg(Faturamento=("Preço Total", "sum"))
        fig = px.pie(canal, names="Loja virtual", values="Faturamento", hole=0.64,
                     color_discrete_sequence=PALETA)
        fig.update_traces(textposition="inside", textinfo="percent", insidetextorientation="horizontal")
        fig.update_layout(legend=dict(orientation="h", x=0.5, xanchor="center", y=-0.02))
        centro_rosca(fig, brl(ind["faturamento"]), "faturamento")
        grafico(fig, 360)

    esq, dir_ = st.columns([1.3, 1])
    with esq:
        st.subheader("Faturamento acumulado" + (" e meta" if meta > 0 else ""))
        fig = go.Figure()
        fig.add_trace(go.Scatter(
            x=diaria["Data"], y=diaria["Acumulado"], mode="lines", fill="tozeroy",
            name="Acumulado", line=dict(color=PURPLE, width=3),
            fillcolor="rgba(143,125,255,0.14)",
        ))
        if meta > 0:
            fig.add_hline(
                y=meta, line_dash="dash", line_color=AMBER,
                annotation_text=f"Meta {brl(meta)}", annotation_font_color=AMBER,
            )
        fig.update_yaxes(title="Faturamento acumulado", tickprefix="R$ ")
        fig.update_xaxes(tickformat="%d/%m")
        grafico(fig, 320)

    with dir_:
        if meta > 0:
            st.subheader("Meta do período")
            fig = go.Figure(go.Indicator(
                mode="gauge+number", value=ind["faturamento"],
                number=dict(prefix="R$ ", valueformat=",.0f", font=dict(size=34, color=TEXT)),
                gauge=dict(
                    axis=dict(range=[0, max(meta, ind["faturamento"]) * 1.1],
                              tickcolor=MUTED, tickfont=dict(color=MUTED)),
                    bar=dict(color=PURPLE, thickness=0.7),
                    bgcolor="rgba(255,255,255,0.05)", borderwidth=0,
                    threshold=dict(line=dict(color=AMBER, width=3), thickness=0.85, value=meta),
                ),
                title=dict(
                    text=f"{ind['faturamento'] / meta:.0%} da meta".replace(".", ","),
                    font=dict(size=14, color=MUTED),
                ),
            ))
            grafico(fig, 320)
        else:
            st.subheader("Faturamento por dia da semana")
            sem = dados.por_dia_semana(f)
            fig = px.bar(sem, x="Dia da semana", y="Faturamento médio",
                         color_discrete_sequence=[TEAL])
            fig.update_yaxes(tickprefix="R$ ")
            grafico(fig, 320)


# ======================= PRODUTOS =======================
with aba_produtos:
    abc = dados.curva_abc(f)
    st.subheader("Curva ABC: onde está o faturamento")
    if abc.empty:
        st.info("Sem faturamento positivo no recorte.")
    else:
        curto = abc["Descrição"].str.slice(0, 24) + abc["Descrição"].str.len().gt(24).map(
            {True: "…", False: ""}
        )
        abc = abc.assign(Produto=(abc.index + 1).astype(str) + ". " + curto)
        cores = abc["Classe"].map({"A": PURPLE, "B": TEAL, "C": CINZA})
        fig = go.Figure()
        fig.add_trace(go.Bar(
            x=abc["Produto"], y=abc["Faturamento"], marker_color=cores, name="Faturamento",
            customdata=abc[["Descrição", "Classe"]],
            hovertemplate="%{customdata[0]}<br>R$ %{y:,.2f}<br>Classe %{customdata[1]}<extra></extra>",
        ))
        fig.add_trace(go.Scatter(
            x=abc["Produto"], y=abc["% acumulado"], yaxis="y2", mode="lines+markers",
            name="% acumulado", line=dict(color=AMBER, width=2), marker=dict(size=6),
        ))
        fig.update_layout(
            yaxis=dict(title="Faturamento", tickprefix="R$ "),
            yaxis2=dict(overlaying="y", side="right", range=[0, 105],
                        ticksuffix="%", showgrid=False),
            legend=dict(orientation="h", y=1.1, x=0), xaxis=dict(tickangle=-35),
        )
        grafico(fig, 430)

    st.subheader("Desempenho por produto")
    prod = dados.desempenho_produtos(f)
    tabela(
        prod,
        {"Faturamento": brl, "Custo": brl, "Margem": brl, "Margem %": "{:.1%}",
         "Itens": inteiro, "Pedidos": inteiro},
    )


# ======================= CANAIS =======================
with aba_canais:
    st.subheader("Comparativo entre canais")
    ch = f.groupby("Loja virtual", as_index=False).agg(
        Faturamento=("Preço Total", "sum"),
        Pedidos=("N° do Pedido", "nunique"),
        Itens=("Quantidade", "sum"),
    )
    ch["Ticket médio"] = ch["Faturamento"] / ch["Pedidos"].where(ch["Pedidos"] > 0)
    canc = dados.taxa_cancelamento(f_todas, "Loja virtual")[["Loja virtual", "Taxa de cancelamento"]]
    ch = ch.merge(canc, on="Loja virtual", how="left")

    a, b = st.columns(2)
    with a:
        fig = px.bar(
            ch.sort_values("Faturamento", ascending=False), x="Loja virtual", y="Faturamento",
            color="Loja virtual", color_discrete_sequence=PALETA, text_auto=".3s",
            title="Faturamento",
        )
        fig.update_yaxes(tickprefix="R$ ")
        fig.update_layout(showlegend=False)
        grafico(fig, 330)
    with b:
        fig = px.bar(
            ch, x="Loja virtual", y="Taxa de cancelamento", color="Loja virtual",
            color_discrete_sequence=PALETA, title="Taxa de cancelamento",
        )
        fig.update_yaxes(tickformat=".0%")
        fig.update_layout(showlegend=False)
        grafico(fig, 330)

    tabela(
        ch,
        {"Faturamento": brl, "Ticket médio": brl, "Itens": inteiro,
         "Pedidos": inteiro, "Taxa de cancelamento": "{:.1%}"},
    )


# ======================= GEOGRAFIA =======================
with aba_geo:
    geo = f.groupby("Estado", as_index=False).agg(
        Faturamento=("Preço Total", "sum"),
        Pedidos=("N° do Pedido", "nunique"),
        Itens=("Quantidade", "sum"),
    ).sort_values("Faturamento", ascending=False)
    
    geo["Latitude"] = geo["Estado"].map(lambda u: dados.CENTROS_UF.get(u, (None, None))[0])
    geo["Longitude"] = geo["Estado"].map(lambda u: dados.CENTROS_UF.get(u, (None, None))[1])

    st.subheader("Mapa de vendas por estado")
    mapa = geo.dropna(subset=["Latitude", "Longitude"])
    if not mapa.empty:
        fig = px.scatter_geo(
            mapa, lat="Latitude", lon="Longitude", size="Faturamento",
            hover_name="Estado",
            hover_data={"Faturamento": ":,.2f", "Pedidos": True, "Latitude": False, "Longitude": False},
            color="Faturamento", color_continuous_scale=ESCALA_ROXA, size_max=45,
        )
        fig.update_geos(
            scope="south america", fitbounds="locations", bgcolor="rgba(0,0,0,0)",
            showframe=False, showcoastlines=False, showocean=False,
            showland=True, landcolor="#18203A", showcountries=True,
            countrycolor="rgba(255,255,255,0.14)", showsubunits=True,
        )
        fig.update_coloraxes(showscale=False)
        grafico(fig, 480)

    tabela(
        geo.drop(columns=["Latitude", "Longitude"]),
        {"Faturamento": brl, "Pedidos": inteiro, "Itens": inteiro},
    )


# ======================= DADOS E EXPORTAÇÃO =======================
with aba_dados:
    st.subheader("Registros filtrados")
    colunas = [col for col in dados.COLUNAS_OBRIGATORIAS if col in f.columns]
    saida = f[colunas].copy()
    saida["Quantidade"] = saida["Quantidade"].round().astype("Int64")
    saida["Nome do contato"] = saida["Nome do contato"].map(dados.mascarar_nome)

    tabela(saida, None, 440)
    csv_saida = saida.to_csv(
        index=False, sep=";", decimal=",", date_format="%d/%m/%Y"
    ).encode("utf-8-sig")
    st.download_button(
        "Baixar recorte filtrado (CSV)", data=csv_saida,
        file_name="recorte_vendas_filtrado.csv", mime="text/csv",
    )

st.markdown(
    "<div class='vt-footer'><span>Chypps Tecnologia</span>"
    "<span>Base original utilizada somente para leitura.</span></div>",
    unsafe_allow_html=True,
)