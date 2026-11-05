"""Camada de dados do dashboard Chypps Tecnologia.

Consome a base em Excel (.xlsx) ou CSV somente para leitura, mantendo
a compatibilidade com as métricas e medidas DAX do projeto.
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
    "Faturamento Bruto", "Faturamento Liquido",
]
NUMERICAS = [
    "Desconto pedido", "Frete", "Quantidade", "Valor unitário", 
    "Preço de Custo", "Preço Total", "Faturamento Bruto", "Faturamento Liquido"
]
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
    colunas_origem = {
        "N.º de venda", "Estado", "Data de venda", "Cidade", "UF",
        "Título do anúncio", "Unidades", "Custo do item", "Total (BRL)",
        "Conta", "SKU", "Receita por envio (BRL)",
        "Valor de desconto do carrinho", "NF-e em anexo",
        "Receita por produtos (BRL)"
    }
    if not colunas_origem.issubset(set(bruto.columns)):
        return bruto

    df = pd.DataFrame(index=bruto.index)
    df["ID"] = [f"L{i:04d}" for i in range(1, len(bruto) + 1)]
    df["N° do Pedido"] = bruto["N.º de venda"]
    df["Data"] = bruto["Data de venda"]
    df["Município"] = bruto["Cidade"]
    df["Estado"] = bruto["UF"]
    
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
    df["Situação"] = bruto["Estado"]
    df["Descrição"] = bruto["Título do anúncio"]
    df["Quantidade"] = bruto["Unidades"]
    df["Valor unitário"] = bruto.get("Preço unitário de venda do anúncio (BRL)", pd.Series(index=bruto.index, dtype="string"))
    df["Preço de Custo"] = bruto["Custo do item"]
    
    df["Faturamento Bruto"] = bruto["Receita por produtos (BRL)"]
    df["Faturamento Liquido"] = bruto["Total (BRL)"]
    df["Preço Total"] = bruto["Total (BRL)"]
    
    df["Loja virtual"] = bruto["Conta"]
    df["Código do produto"] = bruto["SKU"]
    df["Frete"] = bruto["Receita por envio (BRL)"]
    df["Desconto pedido"] = bruto["Valor de desconto do carrinho"]
    df["Situação da NFe"] = bruto["NF-e em anexo"].replace({"Não emitida": "Pendente"})
    df["Nome do contato"] = bruto.get("Comprador", bruto.get("Nome do destinatário", pd.Series(index=bruto.index, dtype="string")))
    df["Bairro"] = bruto.get("Bairro", pd.Series(index=bruto.index, dtype="string"))
    return df[COLUNAS_OBRIGATORIAS]


def ler_excel(conteudo: bytes) -> pd.DataFrame:
    try:
        bruto = pd.read_excel(io.BytesIO(conteudo), dtype=str)
    except Exception as erro:
        raise ValueError(f"Não foi possível interpretar o arquivo Excel. Detalhe: {erro}")
    
    bruto.columns = bruto.columns.astype(str).str.strip()
    if "N.º de venda" in bruto.columns:
        return adaptar_exportacao_original(bruto)
    return bruto


def preparar(bruto: pd.DataFrame) -> pd.DataFrame:
    df = bruto.copy()
    for col in NUMERICAS:
        if col in df.columns:
            df[col] = numero_br(df[col])
        else:
            df[col] = 0.0
            
    # Conversão flexível de datas (aceita ISO YYYY-MM-DD e BR DD/MM/YYYY)
    if pd.api.types.is_datetime64_any_dtype(df["Data"]):
        datas = df["Data"]
    else:
        texto = df["Data"].astype("string").str.strip()
        datas = pd.to_datetime(texto, errors="coerce")
        resto = datas.isna() & texto.notna()
        if resto.any():
            datas = datas.where(~resto, pd.to_datetime(texto[resto], dayfirst=True, errors="coerce"))
    
    df["Data"] = datas
    for col in df.columns:
        if col not in NUMERICAS and col != "Data":
            df[col] = df[col].astype("string").str.strip()

    df["Mês"] = df["Data"].dt.to_period("M").astype("string")
    df["Dia"] = df["Data"].dt.day
    df["Dia da semana"] = df["Data"].dt.dayofweek.map(dict(enumerate(DIAS_SEMANA)))
    df["Custo informado"] = df["Preço de Custo"].fillna(0) > 0
    df["Custo total estimado"] = df["Preço de Custo"].fillna(0) * df["Quantidade"].fillna(0)
    df["Margem bruta estimada"] = df["Faturamento Liquido"].fillna(0) - df["Custo total estimado"]
    df["Linha idêntica"] = bruto.duplicated(keep="first").to_numpy()
    df["Cancelado"] = df["Situação"].fillna("").str.casefold().eq("cancelado")
    return df


def carregar(conteudo: bytes, extensao: str = "xlsx") -> pd.DataFrame:
    if extensao.lower() in ("xlsx", "xls"):
        return preparar(ler_excel(conteudo))
    raise ValueError("Formato não suportado. Utilize .xlsx")


def filtrar(
    df: pd.DataFrame,
    inicio=None,
    fim=None,
    situacoes: Optional[Iterable[str]] = None,
    canais: Optional[Iterable[str]] = None,
    estados: Optional[Iterable[str]] = None,
    produtos: Optional[Iterable[str]] = None,
) -> pd.DataFrame:
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
                                     "Linhas", "Itens", "Valor", "Frete", "Desconto"])
    return df.groupby("N° do Pedido", as_index=False).agg(
        Data=("Data", "min"), Canal=("Loja virtual", "first"), Situação=("Situação", "first"),
        NFe=("Situação da NFe", "first"), Município=("Município", "first"), Estado=("Estado", "first"),
        Linhas=("ID", "size"), Itens=("Quantidade", "sum"), Valor=("Faturamento Liquido", "sum"),
        Frete=("Frete", "first"), Desconto=("Desconto pedido", "first"),
    )


def indicadores(df: pd.DataFrame) -> dict:
    faturamento_liquido = float(df["Faturamento Liquido"].fillna(0).sum())
    faturamento_bruto = float(df["Faturamento Bruto"].fillna(0).sum())
    pedidos = int(df["N° do Pedido"].nunique())
    itens = float(df["Quantidade"].fillna(0).sum())
    po = nivel_pedido(df)
    com_custo = df[df["Custo informado"]]
    fat_com_custo = float(com_custo["Faturamento Liquido"].fillna(0).sum())
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


def calcular_metricas_mes_atual_anterior(df_base: pd.DataFrame, canais=None, estados=None):
    if df_base.empty or df_base["Data"].dropna().empty:
        return 0.0, 0.0, 0.0, 0.0, 0.0

    df_contexto = df_base.copy()
    if canais:
        df_contexto = df_contexto[df_contexto["Loja virtual"].isin(list(canais))]
    if estados:
        df_contexto = df_contexto[df_contexto["Estado"].isin(list(estados))]

    data_maxima = df_contexto["Data"].max()
    if pd.isna(data_maxima):
        return 0.0, 0.0, 0.0, 0.0, 0.0

    ano_atual = data_maxima.year
    mes_atual = data_maxima.month

    mes_anterior_data = data_maxima - pd.DateOffset(months=1)
    ano_alvo = mes_anterior_data.year
    mes_alvo = mes_anterior_data.month

    mask_atual = (df_contexto["Data"].dt.year == ano_atual) & (df_contexto["Data"].dt.month == mes_atual)
    fat_liq_atual = float(df_contexto.loc[mask_atual, "Faturamento Liquido"].sum())
    fat_bruto_atual = float(df_contexto.loc[mask_atual, "Faturamento Bruto"].sum())

    mask_anterior = (df_contexto["Data"].dt.year == ano_alvo) & (df_contexto["Data"].dt.month == mes_alvo)
    fat_liq_ant = float(df_contexto.loc[mask_anterior, "Faturamento Liquido"].sum())
    fat_bruto_ant = float(df_contexto.loc[mask_anterior, "Faturamento Bruto"].sum())

    variacao_mom = (fat_liq_atual - fat_liq_ant) / fat_liq_ant if fat_liq_ant != 0 else 0.0

    return fat_liq_atual, fat_liq_ant, fat_bruto_atual, fat_bruto_ant, variacao_mom


def taxa_cancelamento(df: pd.DataFrame, por: Optional[str] = None):
    po = df.groupby("N° do Pedido