"""Camada de dados do dashboard Chypps Tecnologia.

Tudo aqui é pandas puro (sem Streamlit), para poder ser testado.
A planilha de origem (fVendas.xlsx) é apenas LIDA: nenhuma função grava, sobrescreve
ou altera o arquivo.

Correspondência com as medidas DAX do Power BI
----------------------------------------------
Faturamento Bruto   = SUM(fVendas[Receita por produtos (BRL)])  -> coluna "Receita Produtos"
Faturamento Líquido = SUM(fVendas[Total (BRL)])                 -> coluna "Preço Total"
Mês atual / anterior e % MoM -> ``metricas_mom`` (mesma lógica do ALLEXCEPT: só canal e
situação restringem o cálculo; período, UF e produto são ignorados).
"""
from __future__ import annotations

import io
import unicodedata
from datetime import date, datetime
from typing import Iterable, Optional

import numpy as np
import pandas as pd

COLUNAS_OBRIGATORIAS = [
    "ID", "N° do Pedido", "Data", "Município", "Estado", "Situação",
    "Descrição", "Quantidade", "Valor unitário", "Preço de Custo",
    "Preço Total", "Receita Produtos", "Loja virtual", "Código do produto", "Frete",
    "Desconto pedido", "Situação da NFe", "Nome do contato", "Bairro",
]
NUMERICAS = ["Desconto pedido", "Frete", "Quantidade", "Valor unitário", "Preço de Custo",
             "Preço Total", "Receita Produtos"]
# Colunas da exportação original sem as quais não dá para calcular as medidas.
ORIGEM_OBRIGATORIA = ["Data de venda", "Conta", "Total (BRL)", "Receita por produtos (BRL)"]
DIAS_SEMANA = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"]

GRUPOS_SITUACAO = ["Concluído", "Em andamento", "Cancelado", "Devolvido/reembolsado"]

# Centro aproximado de cada UF (latitude, longitude), usado só para o mapa de bolhas.
CENTROS_UF = {
    "AC": (-9.0, -70.5), "AL": (-9.6, -36.8), "AP": (1.4, -51.8), "AM": (-3.9, -63.9),
    "BA": (-12.5, -41.7), "CE": (-5.2, -39.3), "DF": (-15.8, -47.9), "ES": (-19.6, -40.6),
    "GO": (-15.9, -49.8), "MA": (-5.0, -45.3), "MT": (-12.9, -55.9), "MS": (-20.5, -54.5),
    "MG": (-18.5, -44.6), "PA": (-3.8, -52.5), "PB": (-7.1, -36.8), "PR": (-24.6, -51.5),
    "PE": (-8.4, -37.9), "PI": (-7.7, -42.7), "RJ": (-22.2, -42.6), "RN": (-5.8, -36.6),
    "RS": (-29.7, -53.3), "RO": (-10.9, -62.8), "RR": (2.1, -61.4), "SC": (-27.2, -50.4),
    "SP": (-22.2, -48.8), "SE": (-10.6, -37.4), "TO": (-10.2, -48.3),
}

_NOMES_UF = {
    "acre": "AC", "alagoas": "AL", "amapa": "AP", "amazonas": "AM", "bahia": "BA", "ceara": "CE",
    "distrito federal": "DF", "espirito santo": "ES", "goias": "GO", "maranhao": "MA",
    "mato grosso": "MT", "mato grosso do sul": "MS", "minas gerais": "MG", "para": "PA",
    "paraiba": "PB", "parana": "PR", "pernambuco": "PE", "piaui": "PI", "rio de janeiro": "RJ",
    "rio grande do norte": "RN", "rio grande do sul": "RS", "rondonia": "RO", "roraima": "RR",
    "santa catarina": "SC", "sao paulo": "SP", "sergipe": "SE", "tocantins": "TO",
}

NAO_INFORMADO = "Não informado"


# ----------------------------------------------------------------------------
# Formatação (padrão brasileiro)
# ----------------------------------------------------------------------------
def brl(valor) -> str:
    try:
        valor = float(valor)
    except (TypeError, ValueError):
        valor = 0.0
    if np.isnan(valor):
        return "—"
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def inteiro(valor) -> str:
    try:
        return f"{int(round(float(valor))):,}".replace(",", ".")
    except (TypeError, ValueError):
        return "0"


def pct(valor, casas: int = 1) -> str:
    try:
        valor = float(valor)
    except (TypeError, ValueError):
        return "—"
    if np.isnan(valor):
        return "—"
    return f"{valor * 100:.{casas}f}%".replace(".", ",")


def rotulo_mom(variacao) -> str:
    """Texto da variação mensal, como na medida DAX: 📈 / 📉 / ⚪ + percentual com sinal."""
    try:
        v = float(variacao)
    except (TypeError, ValueError):
        v = 0.0
    if np.isnan(v) or round(v * 100, 2) == 0:
        return "⚪ 0,00%"
    numero = f"{abs(v) * 100:.2f}".replace(".", ",")
    return f"📈 +{numero}%" if v > 0 else f"📉 −{numero}%"


# ----------------------------------------------------------------------------
# Utilidades de limpeza
# ----------------------------------------------------------------------------
def numero_br(serie: pd.Series) -> pd.Series:
    """Converte texto no padrão brasileiro ('1.234,56', 'R$ 94,9') em número.

    Séries que já são numéricas passam direto.
    """
    if pd.api.types.is_numeric_dtype(serie):
        return pd.to_numeric(serie, errors="coerce").astype("float64")
    s = serie.astype("string").str.strip()
    s = s.str.replace(r"R\$\s*", "", regex=True).str.replace(" ", "", regex=False)
    tem_virgula = s.str.contains(",", na=False)
    s = s.where(~tem_virgula, s.str.replace(".", "", regex=False).str.replace(",", ".", regex=False))
    # float64 comum (NaN), e não o tipo anulável do pandas (pd.NA), que não funciona em np.where/np.isnan
    return pd.to_numeric(s, errors="coerce").astype("float64")


def _sem_acento(texto: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFD", texto) if unicodedata.category(c) != "Mn")


def _texto(serie: pd.Series, padrao: Optional[str] = None) -> pd.Series:
    """Texto limpo (object). Vazio, 'null', 'nan' e 'none' viram ausente (ou `padrao`)."""
    def limpo(x):
        if x is None or (isinstance(x, float) and np.isnan(x)) or x is pd.NaT or x is pd.NA:
            return np.nan if padrao is None else padrao
        t = str(x).strip()
        if t.casefold() in ("", "null", "nan", "none", "nat"):
            return np.nan if padrao is None else padrao
        return t
    return pd.Series([limpo(x) for x in serie], index=serie.index, dtype="object")


def normalizar_uf(serie: pd.Series) -> pd.Series:
    """'São Paulo', 'sao paulo' e 'SP' viram 'SP'. Ausente vira 'Não informado'."""
    def um(valor):
        if valor is None or (isinstance(valor, float) and np.isnan(valor)):
            return NAO_INFORMADO
        t = str(valor).strip()
        if t.casefold() in ("", "null", "nan", "none"):
            return NAO_INFORMADO
        if len(t) == 2 and t.upper() in CENTROS_UF:
            return t.upper()
        return _NOMES_UF.get(_sem_acento(t).casefold(), t)
    return serie.map(um).astype("object")


def grupo_situacao(valor) -> str:
    """Agrupa os 30+ textos de status das plataformas em quatro grupos."""
    if valor is None or (isinstance(valor, float) and np.isnan(valor)):
        return "Em andamento"
    t = _sem_acento(str(valor)).casefold()
    if "pode pedir uma devolucao" in t:          # prazo aberto, a venda segue valendo
        return "Concluído"
    if "mediacao finalizada" in t and "reembols" not in t:   # 'Te demos o dinheiro'
        return "Concluído"
    if any(p in t for p in ("reembols", "devolu", "vamos te devolver", "devolvido")):
        return "Devolvido/reembolsado"
    if any(p in t for p in ("cancel", "malsucedido", "nao entregue")):
        return "Cancelado"
    if "conclu" in t:
        return "Concluído"
    return "Em andamento"


def parse_datas(serie: pd.Series) -> pd.Series:
    """Datas reais do Excel ou texto. Inválidas viram NaT.

    Texto 'aaaa-mm-dd ...' é lido como ISO; qualquer outro texto como 'dd/mm/aaaa [hh:mm:ss]'
    (dia primeiro). Ler ISO com dayfirst trocaria dia e mês, por isso os dois casos são separados.
    """
    if pd.api.types.is_datetime64_any_dtype(serie):
        return serie
    objetos = serie.astype("object")
    saida = pd.Series(pd.NaT, index=serie.index, dtype="datetime64[ns]")
    ja_data = objetos.map(lambda x: isinstance(x, (pd.Timestamp, datetime, date)))
    if ja_data.any():
        saida[ja_data] = pd.to_datetime(objetos[ja_data].tolist(), errors="coerce")
    texto = objetos[~ja_data].map(lambda x: np.nan if x is None else str(x).strip())
    iso = texto.str.match(r"^\d{4}-\d{2}-\d{2}", na=False)
    if iso.any():
        saida[iso[iso].index] = pd.to_datetime(texto[iso], errors="coerce").to_numpy()
    resto = texto[~iso & texto.notna()]
    if len(resto):
        saida[resto.index] = pd.to_datetime(resto, dayfirst=True, errors="coerce", format="mixed").to_numpy()
    return saida


# ----------------------------------------------------------------------------
# Leitura e adaptação da exportação original
# ----------------------------------------------------------------------------
_NULOS = ["null", "NULL", "Null"]


def ler_excel(conteudo: bytes) -> pd.DataFrame:
    """Lê a primeira aba do Excel a partir de bytes (somente leitura)."""
    try:
        bruto = pd.read_excel(io.BytesIO(conteudo), na_values=_NULOS)
    except Exception as erro:  # noqa: BLE001
        raise ValueError(f"Não foi possível ler o Excel: {erro}") from erro
    bruto.columns = bruto.columns.astype(str).str.strip()
    return bruto


def ler_csv(conteudo: bytes) -> pd.DataFrame:
    """Lê o CSV a partir de bytes. Aceita ';' ou ',' e UTF-8 ou Latin-1."""
    ultimo: Optional[Exception] = None
    for codificacao in ("utf-8-sig", "latin-1"):
        for sep in (";", ","):
            try:
                bruto = pd.read_csv(io.BytesIO(conteudo), sep=sep, encoding=codificacao,
                                    na_values=_NULOS, decimal="," if sep == ";" else ".")
            except Exception as erro:  # noqa: BLE001
                ultimo = erro
                continue
            bruto.columns = bruto.columns.astype(str).str.strip()
            if "Data de venda" in bruto.columns:
                return bruto
    raise ValueError(f"Não foi possível ler o CSV (colunas da exportação não encontradas). Detalhe: {ultimo}")


def _col(bruto: pd.DataFrame, nome: str) -> pd.Series:
    """Coluna da origem, ou série vazia se a plataforma não a exporta."""
    if nome in bruto.columns:
        return bruto[nome]
    return pd.Series(np.nan, index=bruto.index, dtype="object")


def _iniciais(canal: str) -> str:
    letras = "".join(p[0] for p in _sem_acento(str(canal)).split() if p and p[0].isalnum()).upper()
    return letras or "PD"


def _chaves_pedido(bruto: pd.DataFrame, data_hora: pd.Series, canal: pd.Series, comprador: pd.Series) -> pd.Series:
    """Número do pedido por linha.

    O 'N.º de venda' do Mercado Livre perde precisão no Excel (muitas linhas ficam com o mesmo
    número arredondado, ex.: 2000020000000000). Quando um mesmo número aparece em mais de 30% das linhas de
    um canal, o pedido é reconstruído por canal + data/hora da venda + comprador.
    """
    numero = _texto(_col(bruto, "N.º de venda"))
    numero = numero.map(lambda x: x[:-2] if isinstance(x, str) and x.endswith(".0") else x)
    saida = pd.Series(index=bruto.index, dtype="object")
    for c in canal.unique():
        m = canal == c
        n = int(m.sum())
        maior_parte = numero[m].value_counts().max() / n if n else 0.0
        confiavel = bool(numero[m].notna().all()) and (n < 4 or maior_parte <= 0.3)
        if confiavel:
            saida[m] = numero[m]
        else:
            comp = (c + "|" + data_hora[m].astype(str) + "|" + comprador[m].fillna("").astype(str))
            codigos = pd.factorize(comp)[0] + 1
            saida[m] = [f"{_iniciais(c)}-{k:04d}" for k in codigos]
    # mesmo número em canais diferentes seria outro pedido: prefixa o canal só nesses casos
    por_canal = saida.groupby(saida).transform(lambda s: canal[s.index].nunique())
    colisao = por_canal > 1
    saida[colisao] = canal[colisao] + " " + saida[colisao]
    return saida


def adaptar_exportacao_original(bruto: pd.DataFrame) -> pd.DataFrame:
    """Converte as colunas da exportação (Mercado Livre, Shopee, WooCommerce) no esquema interno."""
    faltando = [c for c in ORIGEM_OBRIGATORIA if c not in bruto.columns]
    if faltando:
        raise ValueError("Colunas obrigatórias ausentes na planilha: " + ", ".join(faltando))

    bruto = bruto.dropna(how="all").copy()
    data_hora = parse_datas(bruto["Data de venda"])
    canal = _texto(bruto["Conta"], "Sem canal")
    comprador = _texto(_col(bruto, "Comprador"))
    nome_entrega = (_texto(_col(bruto, "Primeiro nome (entrega)"), "") + " "
                    + _texto(_col(bruto, "Sobrenome (entrega)"), "")).str.strip()
    comprador = comprador.where(comprador.notna(), nome_entrega.where(nome_entrega != "", None))

    df = pd.DataFrame(index=bruto.index)
    df["ID"] = [f"L{i:04d}" for i in range(1, len(bruto) + 1)]
    df["N° do Pedido"] = _chaves_pedido(bruto, data_hora, canal, comprador)
    df["Data"] = data_hora.dt.normalize()
    df["Data e hora"] = data_hora
    df["Município"] = _texto(_col(bruto, "Cidade"), NAO_INFORMADO)
    df["Estado"] = normalizar_uf(_col(bruto, "UF"))
    df["Situação"] = _texto(_col(bruto, "Estado"), "Sem situação")
    df["Grupo de situação"] = df["Situação"].map(grupo_situacao)
    df["Descrição"] = _texto(_col(bruto, "Título do anúncio"), "(sem descrição)")
    df["Quantidade"] = _col(bruto, "Unidades")
    df["Valor unitário"] = _col(bruto, "Preço unitário de venda do anúncio (BRL)")
    df["Preço de Custo"] = _col(bruto, "Custo do item")
    df["Preço Total"] = bruto["Total (BRL)"]            # Faturamento Líquido (DAX)
    df["Receita Produtos"] = bruto["Receita por produtos (BRL)"]   # Faturamento Bruto (DAX)
    df["Loja virtual"] = canal
    df["Código do produto"] = _texto(_col(bruto, "SKU"), "(sem SKU)")
    df["Frete"] = _col(bruto, "Receita por envio (BRL)")
    df["Desconto pedido"] = _col(bruto, "Valor de desconto do carrinho")
    nfe = _texto(_col(bruto, "NF-e em anexo"), "Sem informação")
    df["Situação da NFe"] = nfe.replace({"Não emitida": "Pendente", "NF-e informada": "Emitida"})
    df["Nome do contato"] = comprador.where(comprador.notna(), "").astype("object")
    df["Bairro"] = _texto(_col(bruto, "Bairro"), NAO_INFORMADO)
    return df


def preparar(bruto: pd.DataFrame) -> pd.DataFrame:
    """Esquema interno + conversões + colunas derivadas. Devolve uma cópia."""
    df = adaptar_exportacao_original(bruto)
    df = df[df["Data"].notna()].copy()          # linha sem data válida não entra em nenhuma medida
    for col in NUMERICAS:
        df[col] = numero_br(df[col])

    df["Mês"] = df["Data"].dt.to_period("M").astype(str)
    df["Dia"] = df["Data"].dt.day
    df["Dia da semana"] = df["Data"].dt.dayofweek.map(dict(enumerate(DIAS_SEMANA)))
    df["Custo informado"] = df["Preço de Custo"].fillna(0) > 0
    df["Custo total estimado"] = df["Preço de Custo"].fillna(0) * df["Quantidade"].fillna(0)
    # Margem sobre a receita de produtos (sem frete): só faz sentido onde há custo informado.
    df["Margem bruta estimada"] = df["Receita Produtos"].fillna(0) - df["Custo total estimado"]
    # Linha idêntica a outra anterior: possível duplicidade (apenas sinalizada, nunca removida).
    df["Linha idêntica"] = df.drop(columns=["ID"]).duplicated(keep="first").to_numpy()
    df["Cancelado"] = df["Grupo de situação"].isin(["Cancelado", "Devolvido/reembolsado"])
    return df.reset_index(drop=True)


def carregar(conteudo: bytes, extensao: str = "xlsx") -> pd.DataFrame:
    bruto = ler_excel(conteudo) if extensao.lower() in ("xlsx", "xls", "xlsm") else ler_csv(conteudo)
    return preparar(bruto)


# ----------------------------------------------------------------------------
# Filtros
# ----------------------------------------------------------------------------
def filtrar(
    df: pd.DataFrame,
    inicio=None,
    fim=None,
    situacoes: Optional[Iterable[str]] = None,
    canais: Optional[Iterable[str]] = None,
    estados: Optional[Iterable[str]] = None,
    produtos: Optional[Iterable[str]] = None,
) -> pd.DataFrame:
    """Aplica filtros. `None` (ou lista vazia, no caso de produtos) significa 'sem filtro'.

    `situacoes` refere-se ao "Grupo de situação" (Concluído, Em andamento, Cancelado, ...).
    """
    mascara = pd.Series(True, index=df.index)
    if inicio is not None and fim is not None:
        mascara &= df["Data"].dt.date.between(inicio, fim)
    if situacoes is not None:
        mascara &= df["Grupo de situação"].isin(list(situacoes))
    if canais is not None:
        mascara &= df["Loja virtual"].isin(list(canais))
    if estados is not None:
        mascara &= df["Estado"].isin(list(estados))
    if produtos:
        mascara &= df["Descrição"].isin(list(produtos))
    return df[mascara].copy()


# ----------------------------------------------------------------------------
# Medidas DAX: faturamento bruto/líquido, mês atual/anterior e % MoM
# ----------------------------------------------------------------------------
def faturamento_bruto(df: pd.DataFrame) -> float:
    """Faturamento Bruto = SUM(Receita por produtos (BRL))."""
    return float(df["Receita Produtos"].sum())


def faturamento_liquido(df: pd.DataFrame) -> float:
    """Faturamento Líquido = SUM(Total (BRL))."""
    return float(df["Preço Total"].sum())


def _variacao(atual: float, anterior: float) -> float:
    """DIVIDE((atual - anterior), anterior, 0)."""
    return (atual - anterior) / anterior if anterior else 0.0


def metricas_mom(df: pd.DataFrame, canais: Optional[Iterable[str]] = None,
                 situacoes: Optional[Iterable[str]] = None) -> dict:
    """Reproduz as medidas 'Mês Atual', 'Mês Anterior' e '% MoM' do Power BI.

    Como no ALLEXCEPT(fVendas, Conta, Estado), só o canal e a situação restringem o cálculo;
    período, UF e produto são ignorados. O mês atual é o mês da última venda nesse recorte e o
    anterior é o mês imediatamente antes (EDATE -1).
    """
    base = filtrar(df, situacoes=situacoes, canais=canais)
    zero = {"mes_atual": None, "mes_anterior": None, "liquido_atual": 0.0, "liquido_anterior": 0.0,
            "bruto_atual": 0.0, "bruto_anterior": 0.0, "var_liquido": 0.0, "var_bruto": 0.0,
            "rotulo_liquido": rotulo_mom(0), "rotulo_bruto": rotulo_mom(0), "tem_anterior": False}
    if base.empty:
        return zero
    maxima = base["Data"].max()
    ant = maxima - pd.DateOffset(months=1)
    cur = base[(base["Data"].dt.year == maxima.year) & (base["Data"].dt.month == maxima.month)]
    pre = base[(base["Data"].dt.year == ant.year) & (base["Data"].dt.month == ant.month)]
    liq_a, liq_p = faturamento_liquido(cur), faturamento_liquido(pre)
    bru_a, bru_p = faturamento_bruto(cur), faturamento_bruto(pre)
    v_liq, v_bru = _variacao(liq_a, liq_p), _variacao(bru_a, bru_p)
    return {
        "mes_atual": f"{maxima:%Y-%m}", "mes_anterior": f"{ant:%Y-%m}",
        "liquido_atual": liq_a, "liquido_anterior": liq_p,
        "bruto_atual": bru_a, "bruto_anterior": bru_p,
        "var_liquido": v_liq, "var_bruto": v_bru,
        "rotulo_liquido": rotulo_mom(v_liq), "rotulo_bruto": rotulo_mom(v_bru),
        "tem_anterior": not pre.empty,
    }


def calcular_metricas_mes_atual_anterior(df: pd.DataFrame, canais=None, estados=None):
    """Compatibilidade com a versão anterior: (liq_atual, liq_ant, bruto_atual, bruto_ant, variação).

    O parâmetro `estados` é ignorado de propósito (o DAX usa ALLEXCEPT e só mantém canal e situação).
    """
    m = metricas_mom(df, canais=canais)
    return (m["liquido_atual"], m["liquido_anterior"], m["bruto_atual"], m["bruto_anterior"], m["var_liquido"])


def resumo_mensal(df: pd.DataFrame) -> pd.DataFrame:
    """Bruto, líquido e pedidos por mês, com variação mensal (% MoM) do líquido e do bruto."""
    if df.empty:
        return pd.DataFrame(columns=["Mês", "Faturamento bruto", "Faturamento líquido", "Pedidos",
                                     "MoM líquido", "MoM bruto"])
    m = df.groupby("Mês", as_index=False).agg(**{
        "Faturamento bruto": ("Receita Produtos", "sum"),
        "Faturamento líquido": ("Preço Total", "sum"),
        "Pedidos": ("N° do Pedido", "nunique"),
    }).sort_values("Mês").reset_index(drop=True)
    m["MoM líquido"] = [np.nan] + [_variacao(a, p) for p, a in zip(m["Faturamento líquido"][:-1], m["Faturamento líquido"][1:])]
    m["MoM bruto"] = [np.nan] + [_variacao(a, p) for p, a in zip(m["Faturamento bruto"][:-1], m["Faturamento bruto"][1:])]
    return m


# ----------------------------------------------------------------------------
# Indicadores
# ----------------------------------------------------------------------------
def nivel_pedido(df: pd.DataFrame) -> pd.DataFrame:
    """Uma linha por pedido. Frete e desconto são do PEDIDO (repetidos em cada item na base),
    por isso entram uma única vez por pedido e nunca são somados linha a linha."""
    if df.empty:
        return pd.DataFrame(columns=["N° do Pedido", "Data", "Canal", "Situação", "NFe", "Município", "Estado",
                                     "Linhas", "Itens", "Valor", "Receita", "Frete", "Desconto"])
    return df.groupby("N° do Pedido", as_index=False).agg(
        Data=("Data", "min"), Canal=("Loja virtual", "first"), Situação=("Grupo de situação", "first"),
        NFe=("Situação da NFe", "first"), Município=("Município", "first"), Estado=("Estado", "first"),
        Linhas=("ID", "size"), Itens=("Quantidade", "sum"), Valor=("Preço Total", "sum"),
        Receita=("Receita Produtos", "sum"), Frete=("Frete", "first"), Desconto=("Desconto pedido", "first"),
    )


def indicadores(df: pd.DataFrame) -> dict:
    liquido = faturamento_liquido(df)
    bruto = faturamento_bruto(df)
    pedidos = int(df["N° do Pedido"].nunique())
    itens = float(df["Quantidade"].fillna(0).sum())
    po = nivel_pedido(df)
    com_custo = df[df["Custo informado"]]
    receita_com_custo = float(com_custo["Receita Produtos"].fillna(0).sum())
    margem = float(com_custo["Margem bruta estimada"].sum())
    frete_informado = po["Frete"].dropna() if pedidos else pd.Series(dtype=float)
    return {
        "faturamento": liquido,                 # mantido: faturamento = líquido
        "faturamento_liquido": liquido,
        "faturamento_bruto": bruto,
        "pedidos": pedidos,
        "itens": itens,
        "ticket_medio": liquido / pedidos if pedidos else 0.0,
        "ticket_medio_bruto": bruto / pedidos if pedidos else 0.0,
        "itens_por_pedido": itens / pedidos if pedidos else 0.0,
        "frete_total": float(po["Frete"].fillna(0).sum()) if pedidos else 0.0,
        "desconto_total": float(po["Desconto"].fillna(0).sum()) if pedidos else 0.0,
        # só entre pedidos com frete informado (em alguns canais a coluna vem vazia)
        "pct_frete_gratis": float((frete_informado == 0).mean()) if len(frete_informado) else np.nan,
        "custo_estimado": float(com_custo["Custo total estimado"].sum()),
        "margem_bruta": margem,
        "margem_pct": margem / receita_com_custo if receita_com_custo else np.nan,
        "cobertura_custo": receita_com_custo / bruto if bruto else np.nan,
        "nfe_pendentes": int((po["NFe"] == "Pendente").sum()) if pedidos else 0,
    }


def taxa_cancelamento(df: pd.DataFrame, por: Optional[str] = None):
    """Pedidos cancelados ou devolvidos / pedidos totais. `df` deve conter TODAS as situações."""
    if df.empty:
        return float("nan") if por is None else pd.DataFrame(columns=[por, "Pedidos", "Cancelados", "Taxa de cancelamento"])
    po = df.groupby("N° do Pedido", as_index=False).agg(
        cancelado=("Cancelado", "max"), grupo=(por, "first") if por else ("Cancelado", "size"))
    if por is None:
        return float(po["cancelado"].mean())
    saida = po.groupby("grupo", as_index=False).agg(Pedidos=("cancelado", "size"), Cancelados=("cancelado", "sum"))
    saida["Taxa de cancelamento"] = saida["Cancelados"] / saida["Pedidos"]
    return saida.rename(columns={"grupo": por})


def serie_diaria(df: pd.DataFrame) -> pd.DataFrame:
    """Faturamento líquido e pedidos por dia, incluindo dias sem venda (zero), com média móvel e acumulado."""
    cols = ["Data", "Faturamento", "Faturamento bruto", "Pedidos", "Média móvel 7 dias", "Acumulado"]
    if df.empty or df["Data"].dropna().empty:
        return pd.DataFrame(columns=cols)
    d = df.groupby("Data").agg(Faturamento=("Preço Total", "sum"), **{"Faturamento bruto": ("Receita Produtos", "sum")},
                               Pedidos=("N° do Pedido", "nunique"))
    d = d.reindex(pd.date_range(d.index.min(), d.index.max(), freq="D"), fill_value=0)
    d.index.name = "Data"
    d["Média móvel 7 dias"] = d["Faturamento"].rolling(7, min_periods=1).mean()
    d["Acumulado"] = d["Faturamento"].cumsum()
    return d.reset_index()[cols]


def por_dia_semana(df: pd.DataFrame) -> pd.DataFrame:
    """Faturamento MÉDIO por dia da semana (média dos dias, para não favorecer dias que aparecem mais vezes)."""
    d = serie_diaria(df)
    if d.empty:
        return pd.DataFrame(columns=["Dia da semana", "Faturamento médio"])
    d["Dia da semana"] = d["Data"].dt.dayofweek.map(dict(enumerate(DIAS_SEMANA)))
    saida = d.groupby("Dia da semana", as_index=False).agg(**{"Faturamento médio": ("Faturamento", "mean")})
    saida["ordem"] = saida["Dia da semana"].map({n: i for i, n in enumerate(DIAS_SEMANA)})
    return saida.sort_values("ordem").drop(columns="ordem")


def desempenho_produtos(df: pd.DataFrame) -> pd.DataFrame:
    cols = ["Código do produto", "Descrição", "Faturamento", "Itens", "Pedidos", "Custo", "Margem", "Margem %"]
    if df.empty:
        return pd.DataFrame(columns=cols)
    p = df.groupby(["Código do produto", "Descrição"], dropna=False, as_index=False).agg(
        Faturamento=("Preço Total", "sum"), Itens=("Quantidade", "sum"), Pedidos=("N° do Pedido", "nunique"),
        Receita=("Receita Produtos", "sum"),
        Custo=("Custo total estimado", "sum"), Margem=("Margem bruta estimada", "sum"),
        Custo_informado=("Custo informado", "all"),
    )
    # Sem custo informado a margem seria 100% (falsa): fica em branco.
    p["Margem"] = p["Margem"].where(p["Custo_informado"])
    p["Margem %"] = np.where((p["Receita"] != 0) & p["Custo_informado"], p["Margem"] / p["Receita"].replace(0, np.nan), np.nan)
    p["Custo"] = p["Custo"].where(p["Custo_informado"])
    return p.drop(columns=["Custo_informado", "Receita"]).sort_values("Faturamento", ascending=False)[cols]


def curva_abc(df: pd.DataFrame) -> pd.DataFrame:
    """Curva ABC (Pareto) por produto: A até 80% do faturamento acumulado, B até 95%, C o restante."""
    p = df.groupby("Descrição", as_index=False).agg(Faturamento=("Preço Total", "sum"))
    p = p[p["Faturamento"] > 0].sort_values("Faturamento", ascending=False).reset_index(drop=True)
    total = p["Faturamento"].sum()
    if p.empty or total == 0:
        p["% do total"] = pd.Series(dtype=float)
        p["% acumulado"] = pd.Series(dtype=float)
        p["Classe"] = pd.Series(dtype=object)
        return p
    p["% do total"] = p["Faturamento"] / total * 100
    p["% acumulado"] = p["% do total"].cumsum()
    anterior = p["% acumulado"] - p["% do total"]
    p["Classe"] = np.where(anterior < 80, "A", np.where(anterior < 95, "B", "C"))
    return p


def mascarar_nome(nome) -> str:
    """'Ana Júlia Lopes' -> 'A. J. L.' (privacidade por padrão, em linha com a LGPD)."""
    if nome is None or pd.isna(nome):
        return ""
    return " ".join(f"{p[0].upper()}." for p in str(nome).split() if p)


# ----------------------------------------------------------------------------
# Destaques em linguagem simples
# ----------------------------------------------------------------------------
def destaques(df: pd.DataFrame, df_todas_situacoes: Optional[pd.DataFrame] = None) -> list[str]:
    if df.empty:
        return []
    ind = indicadores(df)
    frases: list[str] = []

    diaria = serie_diaria(df)
    if not diaria.empty and diaria["Faturamento"].max() > 0:
        melhor = diaria.loc[diaria["Faturamento"].idxmax()]
        frases.append(f"Melhor dia: {melhor['Data']:%d/%m} ({brl(melhor['Faturamento'])}, {inteiro(melhor['Pedidos'])} pedidos).")

    abc = curva_abc(df)
    if not abc.empty:
        lider = abc.iloc[0]
        n_a = int((abc["Classe"] == "A").sum())
        frases.append(f"Produto líder: {lider['Descrição']} ({pct(lider['% do total'] / 100)} do faturamento). "
                      f"{n_a} de {len(abc)} produtos concentram 80% das vendas (classe A).")

    canal = df.groupby("Loja virtual")["Preço Total"].sum().sort_values(ascending=False)
    if not canal.empty and canal.sum() > 0:
        frases.append(f"Canal principal: {canal.index[0]} ({pct(canal.iloc[0] / canal.sum())} do faturamento líquido).")

    uf = df.groupby("Estado")["Preço Total"].sum().sort_values(ascending=False)
    uf = uf[uf.index != NAO_INFORMADO]
    if not uf.empty and uf.sum() > 0:
        top3 = ", ".join(f"{u} ({pct(v / uf.sum(), 0)})" for u, v in uf.head(3).items())
        frases.append(f"Estados que mais compram: {top3}.")

    if df_todas_situacoes is not None and not df_todas_situacoes.empty:
        t = taxa_cancelamento(df_todas_situacoes)
        frases.append(f"Cancelamentos e devoluções no recorte: {pct(t)} dos pedidos.")

    if not np.isnan(ind["pct_frete_gratis"]):
        frases.append(f"{pct(ind['pct_frete_gratis'])} dos pedidos com frete informado têm frete zero.")
    if ind["nfe_pendentes"]:
        n = ind["nfe_pendentes"]
        frases.append(f"{inteiro(n)} {'pedido está' if n == 1 else 'pedidos estão'} com a nota fiscal pendente.")
    return frases


def qualidade(df_completo: pd.DataFrame) -> dict:
    """Sinais de qualidade da base. Servem para ALERTAR; nenhuma linha é removida."""
    multi = df_completo.groupby("N° do Pedido").agg(n=("ID", "size"), v=("Preço Total", "nunique"))
    return {
        "linhas": int(len(df_completo)),
        "pedidos": int(df_completo["N° do Pedido"].nunique()),
        "datas_invalidas": int(df_completo["Data"].isna().sum()),
        "linhas_identicas": int(df_completo["Linha idêntica"].sum()),
        "linhas_sem_custo": int((~df_completo["Custo informado"]).sum()),
        "pedidos_multilinha": int((multi["n"] > 1).sum()),
        # pedidos com várias linhas e o MESMO total repetido em todas: o SUM do DAX conta o total várias vezes
        "pedidos_total_repetido": int(((multi["n"] > 1) & (multi["v"] == 1)).sum()),
        "uf_nao_informada": int((df_completo["Estado"] == NAO_INFORMADO).sum()),
    }
