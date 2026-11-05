"""Camada de dados do dashboard Chypps Tecnologia.

Tudo aqui é pandas puro (sem Streamlit). A base Excel é apenas LIDA:
nenhuma função grava, sobrescreve ou altera o arquivo de origem.
"""
from __future__ import annotations

import io
from typing import Iterable, Optional

import numpy as np
import pandas as pd

COLUNAS_OBRIGATORIAS = [
    "ID", "N° do Pedido", "Data", "Município", "Estado", "Situação",
    "Descrição", "Quantidade", "Valor unitário", "Preço de Custo",
    "Preço Total", "Loja virtual", "Código do produto", "Frete",
    "Desconto pedido", "Situação da NFe", "Nome do contato", "Bairro",
]
NUMERICAS = ["Desconto pedido", "Frete", "Quantidade", "Valor unitário", "Preço de Custo", "Preço Total", "Faturamento Bruto"]
DIAS_SEMANA = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"]

CENTROS_UF = {
    "AC": (-9.0, -70.5), "AL": (-9.6, -36.8), "AP": (1.4, -51.8), "AM": (-3.9, -63.9),
    "BA": (-12.5, -41.7), "CE": (-5.2, -39.3), "DF": (-15.8, -47.9), "ES": (-19.6, -40.6),
    "GO": (-15.9, -49.8), "MA": (-5.0, -45.3), "MT": (-12.9, -55.9), "MS": (-20.5, -54.5),
    "MG": (-18.5, -44.6), "PA": (-3.8, -52.5), "PB": (-7.1, -36.8), "PR": (-24.6, -51.5),
    "PE": (-8.4, -37.9), "PI": (-7.7, -42.7), "RJ": (-22.2, -42.6), "RN": (-5.8, -36.6),
    "RS": (-29.7, -53.3), "RO": (-10.9, -62.8), "RR": (2.1, -61.4), "SC": (-27.2, -50.4),
    "SP": (-22.2, -48.8), "SE": (-10.6, -37.4), "TO": (-10.2, -48.3),
}


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


def numero_br(serie: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(serie):
        return pd.to_numeric(serie, errors="coerce").astype("float64")
    s = serie.astype("string").str.strip()
    s = s.str.replace(r"R\$\s*", "", regex=True).str.replace(" ", "", regex=False)
    tem_virgula = s.str.contains(",", na=False)
    s = s.where(~tem_virgula, s.str.replace(".", "", regex=False).str.replace(",", ".", regex=False))
    return pd.to_numeric(s, errors="coerce").astype("float64")


def adaptar_exportacao_original(bruto: pd.DataFrame) -> pd.DataFrame:
    df = pd.DataFrame(index=bruto.index)
    df["ID"] = [f"L{i:04d}" for i in range(1, len(bruto) + 1)]
    df["N° do Pedido"] = bruto.get("N.º de venda", bruto.get("N° do Pedido"))
    df["Data"] = bruto.get("Data de venda", bruto.get("Data"))
    df["Município"] = bruto.get("Cidade", bruto.get("Município"))
    df["Estado"] = bruto.get("UF", bruto.get("Estado"))
    
    mapa_uf = {
        "Acre": "AC", "Alagoas": "AL", "Amapá": "AP", "Amazonas": "AM",
        "Bahia": "BA", "Ceará": "CE", "Distrito Federal": "DF", "Espírito Santo": "ES",
        "Goiás": "GO", "Maranhão": "MA", "Mato Grosso": "MT", "Mato Grosso do Sul": "MS",
        "Minas Gerais": "MG", "Pará": "PA", "Paraíba": "PB", "Paraná": "PR",
        "Pernambuco": "PE", "Piauí": "PI", "Rio de Janeiro": "RJ", "Rio Grande do Norte": "RN",
        "Rio Grande do Sul": "RS", "Rondônia": "RO", "Roraima": "RR", "Santa Catarina": "SC",
        "São Paulo": "SP", "Sergipe": "SE", "Tocantins": "TO",
    }
    df["Estado"] = df["Estado"].astype("string").str.strip().replace(mapa_uf).str.upper()
    df["Situação"] = bruto.get("Estado", bruto.get("Situação"))
    df["Descrição"] = bruto.get("Título do anúncio", bruto.get("Descrição"))
    df["Quantidade"] = bruto.get("Unidades", bruto.get("Quantidade"))
    df["Valor unitário"] = bruto.get("Preço unitário de venda do anúncio (BRL)", bruto.get("Valor unitário"))
    df["Preço de Custo"] = bruto.get("Custo do item", bruto.get("Preço de Custo"))
    
    # Faturamento Líquido (Total BRL) e Faturamento Bruto (Receita por produtos BRL)
    df["Preço Total"] = bruto.get("Total (BRL)", bruto.get("Preço Total"))
    df["Faturamento Bruto"] = bruto.get("Receita por produtos (BRL)", df["Preço Total"])
    
    df["Loja virtual"] = bruto.get("Conta", bruto.get("Loja virtual"))
    df["Código do produto"] = bruto.get("SKU", bruto.get("Código do produto"))
    df["Frete"] = bruto.get("Receita por envio (BRL)", bruto.get("Frete"))
    df["Desconto pedido"] = bruto.get("Valor de desconto do carrinho", bruto.get("Desconto pedido"))
    df["Situação da NFe"] = bruto.get("NF-e em anexo", bruto.get("Situação da NFe")).replace({"Não emitida": "Pendente"})
    df["Nome do contato"] = bruto.get("Comprador", bruto.get("Nome do destinatário", bruto.get("Nome do contato")))
    df["Bairro"] = bruto.get("Bairro", pd.Series(index=bruto.index, dtype="string"))
    return df[COLUNAS_OBRIGATORIAS + ["Faturamento Bruto"]]


def ler_excel(conteudo: bytes) -> pd.DataFrame:
    try:
        bruto = pd.read_excel(io.BytesIO(conteudo), dtype=str)
    except Exception as erro:
        raise ValueError(f"Não foi possível interpretar o arquivo Excel. Detalhe: {erro}")
    bruto.columns = bruto.columns.astype(str).str.strip()
    return adaptar_exportacao_original(bruto)


def ler_csv(conteudo: bytes) -> pd.DataFrame:
    ultimo_erro: Optional[Exception] = None
    for codificacao in ("utf-8-sig", "latin-1"):
        for sep in (";", ","):
            try:
                bruto = pd.read_csv(io.BytesIO(conteudo), sep=sep, dtype="string", encoding=codificacao)
            except Exception as erro:
                ultimo_erro = erro
                continue
            bruto.columns = bruto.columns.astype(str).str.strip()
            return adaptar_exportacao_original(bruto)
    raise ValueError(f"Não foi possível interpretar o CSV. Detalhe: {ultimo_erro}")


def preparar(bruto: pd.DataFrame) -> pd.DataFrame:
    df = bruto.copy()
    num_cols = NUMERICAS + ["Faturamento Bruto"]
    for col in num_cols:
        if col in df.columns:
            df[col] = numero_br(df[col])
            
    texto = df["Data"].astype("string").str.strip()
    datas = pd.to_datetime(texto, format="%d/%m/%Y %H:%M:%S", errors="coerce")
    resto = datas.isna() & texto.notna()
    if resto.any():
        datas = datas.where(~resto, pd.to_datetime(texto, format="%d/%m/%Y", dayfirst=True, errors="coerce"))
    df["Data"] = datas
    
    for col in df.columns:
        if col not in num_cols and col != "Data":
            df[col] = df[col].astype("string").str.strip()

    df["Mês"] = df["Data"].dt.to_period("M").astype("string")
    df["Dia"] = df["Data"].dt.day
    df["Dia da semana"] = df["Data"].dt.dayofweek.map(dict(enumerate(DIAS_SEMANA)))
    df["Custo informado"] = df["Preço de Custo"].fillna(0) > 0
    df["Custo total estimado"] = df["Preço de Custo"].fillna(0) * df["Quantidade"].fillna(0)
    df["Margem bruta estimada"] = df["Preço Total"].fillna(0) - df["Custo total estimado"]
    df["Linha idêntica"] = bruto.duplicated(keep="first").to_numpy()
    df["Cancelado"] = df["Situação"].fillna("").str.casefold().eq("cancelado")
    return df


def carregar(conteudo: bytes, extensao: str = "xlsx") -> pd.DataFrame:
    if extensao.lower() in ("xlsx", "xls"):
        return preparar(ler_excel(conteudo))
    return preparar(ler_csv(conteudo))


def filtrar(df: pd.DataFrame, inicio=None, fim=None, situacoes=None, canais=None, estados=None, produtos=None) -> pd.DataFrame:
    mascara = pd.Series(True, index=df.index)
    if inicio is not None and fim is not None:
        datas = df["Data"].dt.date
        mascara &= datas.between(inicio, fim)
    if situacoes is not None:
        mascara &= df["Situação"].isin(list(situacoes))
    if canais is not None:
        mascara &= df["Loja virtual"].isin(list(canais))
    if estados is not None:
        mascara &= df["Estado"].isin(list(estados))
    if produtos:
        mascara &= df["Descrição"].isin(list(produtos))
    return df[mascara].copy()


def nivel_pedido(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["N° do Pedido", "Data", "Canal", "Situação", "NFe", "Município", "Estado",
                                     "Linhas", "Itens", "Valor", "Faturamento Bruto", "Frete", "Desconto"])
    return df.groupby("N° do Pedido", as_index=False).agg(
        Data=("Data", "min"), Canal=("Loja virtual", "first"), Situação=("Situação", "first"),
        NFe=("Situação da NFe", "first"), Município=("Município", "first"), Estado=("Estado", "first"),
        Linhas=("ID", "size"), Itens=("Quantidade", "sum"), Valor=("Preço Total", "sum"),
        Faturamento Bruto=("Faturamento Bruto", "sum"),
        Frete=("Frete", "first"), Desconto=("Desconto pedido", "first"),
    )


def indicadores(df: pd.DataFrame) -> dict:
    faturamento_liquido = float(df["Preço Total"].fillna(0).sum())
    faturamento_bruto = float(df["Faturamento Bruto"].fillna(0).sum()) if "Faturamento Bruto" in df.columns else faturamento_liquido
    pedidos = int(df["N° do Pedido"].nunique())
    itens = float(df["Quantidade"].fillna(0).sum())
    po = nivel_pedido(df)
    com_custo = df[df["Custo informado"]]
    fat_com_custo = float(com_custo["Preço Total"].fillna(0).sum())
    margem = float(com_custo["Margem bruta estimada"].sum())
    
    return {
        "faturamento": faturamento_liquido,
        "faturamento_bruto": faturamento_bruto,
        "pedidos": pedidos,
        "itens": itens,
        "ticket_medio": faturamento_liquido / pedidos if pedidos else 0.0,
        "itens_por_pedido": itens / pedidos if pedidos else 0.0,
        "frete_total": float(po["Frete"].fillna(0).sum()) if pedidos else 0.0,
        "desconto_total": float(po["Desconto"].fillna(0).sum()) if pedidos else 0.0,
        "pct_frete_gratis": float((po["Frete"].fillna(0) == 0).mean()) if pedidos else np.nan,
        "custo_estimado": float(com_custo["Custo total estimado"].sum()),
        "margem_bruta": margem,
        "margem_pct": margem / fat_com_custo if fat_com_custo else np.nan,
        "cobertura_custo": fat_com_custo / faturamento_liquido if faturamento_liquido else np.nan,
        "nfe_pendentes": int((po["NFe"] == "Pendente").sum()) if pedidos else 0,
    }


def taxa_cancelamento(df: pd.DataFrame, por: Optional[str] = None):
    po = df.groupby("N° do Pedido", as_index=False).agg(
        cancelado=("Cancelado", "max"), grupo=(por, "first") if por else ("Cancelado", "size"))
    if por is None:
        return float(po["cancelado"].mean()) if len(po) else np.nan
    saida = po.groupby("grupo", as_index=False).agg(Pedidos=("cancelado", "size"), Cancelados=("cancelado", "sum"))
    saida["Taxa de cancelamento"] = saida["Cancelados"] / saida["Pedidos"]
    return saida.rename(columns={"grupo": por})


def serie_diaria(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty or df["Data"].dropna().empty:
        return pd.DataFrame(columns=["Data", "Faturamento", "Pedidos", "Média móvel 7 dias", "Acumulado"])
    d = df.groupby("Data").agg(Faturamento=("Preço Total", "sum"), Pedidos=("N° do Pedido", "nunique"))
    intervalo = pd.date_range(d.index.min(), d.index.max(), freq="D")
    d = d.reindex(intervalo, fill_value=0)
    d.index.name = "Data"
    d["Média móvel 7 dias"] = d["Faturamento"].rolling(7, min_periods=1).mean()
    d["Acumulado"] = d["Faturamento"].cumsum()
    return d.reset_index()


def curva_abc(df: pd.DataFrame) -> pd.DataFrame:
    p = df.groupby("Descrição", as_index=False).agg(Faturamento=("Preço Total", "sum"))
    p = p[p["Faturamento"] > 0].sort_values("Faturamento", ascending=False).reset_index(drop=True)
    total = p["Faturamento"].sum()
    if p.empty or total == 0:
        p["% do total"] = []
        p["% acumulado"] = []
        p["Classe"] = []
        return p
    p["% do total"] = p["Faturamento"] / total * 100
    p["% acumulado"] = p["% do total"].cumsum()
    anterior = p["% acumulado"] - p["% do total"]
    p["Classe"] = np.where(anterior < 80, "A", np.where(anterior < 95, "B", "C"))
    return p


def mascarar_nome(nome) -> str:
    if nome is None or (isinstance(nome, float) and np.isnan(nome)) or pd.isna(nome):
        return ""
    partes = [p for p in str(nome).split() if p]
    return " ".join(f"{p[0].upper()}." for p in partes)


def destaques(df: pd.DataFrame, df_todas_situacoes: Optional[pd.DataFrame] = None) -> list[str]:
    if df.empty:
        return []
    ind = indicadores(df)
    frases: list[str] = []
    diaria = serie_diaria(df)
    if not diaria.empty and diaria["Faturamento"].max() > 0:
        melhor = diaria.loc[diaria["Faturamento"].idxmax()]
        frases.append(f"Melhor dia: {melhor['Data']:%d/%m} ({brl(melhor['Faturamento'])}, {inteiro(melhor['Pedidos'])} pedidos).")
    return frases