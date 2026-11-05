"""Camada de dados do dashboard Chypps Tecnologia.

Consome o Excel original ou CSV em modo de leitura estrita.
"""
from __future__ import annotations

import io
from typing import Iterable, Optional

import numpy as np
import pandas as pd

COLUNAS_OBRIGATORIAS = [
    "ID", "N° do Pedido", "Data", "Município", "Estado", "Situação",
    "Descrição", "Quantidade", "Valor unitário", "Preço de Custo",
    "Preço Total", "Receita Produtos", "Loja virtual", "Código do produto", "Frete",
    "Desconto pedido", "Situação da NFe", "Nome do contato", "Bairro",
]
NUMERICAS = ["Desconto pedido", "Frete", "Quantidade", "Valor unitário", "Preço de Custo", "Preço Total", "Receita Produtos"]
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
    bruto = bruto.dropna(how="all")
    df = pd.DataFrame(index=bruto.index)
    df["ID"] = [f"L{i:04d}" for i in range(1, len(bruto) + 1)]
    df["N° do Pedido"] = bruto.get("N.º de venda", bruto.get("N° do Pedido", pd.Series(index=bruto.index, dtype="string")))
    df["Data"] = bruto.get("Data de venda", bruto.get("Data", pd.Series(index=bruto.index, dtype="string")))
    df["Município"] = bruto.get("Cidade", bruto.get("Município", pd.Series(index=bruto.index, dtype="string")))
    df["Estado"] = bruto.get("UF", bruto.get("Estado", pd.Series(index=bruto.index, dtype="string")))
    
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
    df["Situação"] = bruto.get("Estado", pd.Series("Concluído", index=bruto.index))
    df["Descrição"] = bruto.get("Título do anúncio", bruto.get("Descrição", pd.Series(index=bruto.index, dtype="string")))
    df["Quantidade"] = bruto.get("Unidades", bruto.get("Quantidade", 1))
    df["Valor unitário"] = bruto.get("Preço unitário de venda do anúncio (BRL)", bruto.get("Valor unitário", 0))
    df["Preço de Custo"] = bruto.get("Custo do item", bruto.get("Preço de Custo", 0))
    df["Preço Total"] = bruto.get("Total (BRL)", bruto.get("Preço Total", 0))
    df["Receita Produtos"] = bruto.get("Receita por produtos (BRL)", df["Preço Total"])
    df["Loja virtual"] = bruto.get("Conta", bruto.get("Loja virtual", "Geral"))
    df["Código do produto"] = bruto.get("SKU", bruto.get("Código do produto", ""))
    df["Frete"] = bruto.get("Receita por envio (BRL)", bruto.get("Frete", 0))
    df["Desconto pedido"] = bruto.get("Valor de desconto do carrinho", bruto.get("Desconto pedido", 0))
    df["Situação da NFe"] = bruto.get("NF-e em anexo", bruto.get("Situação da NFe", "Emitida")).replace({"Não emitida": "Pendente"})
    df["Nome do contato"] = bruto.get("Comprador", bruto.get("Nome do destinatário", bruto.get("Nome do contato", "")))
    df["Bairro"] = bruto.get("Bairro", "")
    return df[COLUNAS_OBRIGATORIAS]


def carregar(conteudo: bytes, extensao: str = "xlsx") -> pd.DataFrame:
    try:
        if extensao.lower() in ("xlsx", "xls"):
            bruto = pd.read_excel(io.BytesIO(conteudo))
        else:
            bruto = None
            for cod in ("utf-8-sig", "latin-1"):
                for sep in (";", ","):
                    try:
                        bruto = pd.read_csv(io.BytesIO(conteudo), sep=sep, encoding=cod)
                        break
                    except Exception:
                        continue
                if bruto is not None:
                    break
            if bruto is None:
                raise ValueError("Não foi possível ler o CSV.")
    except Exception as e:
        raise ValueError(f"Erro ao interpretar arquivo: {e}")

    bruto.columns = bruto.columns.astype(str).str.strip()
    df = adaptar_exportacao_original(bruto)
    df = df.dropna(subset=["N° do Pedido"])

    for col in NUMERICAS:
        if col in df.columns:
            df[col] = numero_br(df[col])

    # Conversão robusta de data sem forçar dayfirst incorreto
    if pd.api.types.is_datetime64_any_dtype(df["Data"]):
        df["Data"] = pd.to_datetime(df["Data"])
    else:
        df["Data"] = pd.to_datetime(df["Data"], errors="coerce")

    df = df.dropna(subset=["Data"])
    df["Mês"] = df["Data"].dt.to_period("M").astype("string")
    df["Dia"] = df["Data"].dt.day
    df["Dia da semana"] = df["Data"].dt.dayofweek.map(dict(enumerate(DIAS_SEMANA)))
    df["Custo informado"] = df["Preço de Custo"].fillna(0) > 0
    df["Custo total estimado"] = df["Preço de Custo"].fillna(0) * df["Quantidade"].fillna(0)
    df["Margem bruta estimada"] = df["Preço Total"].fillna(0) - df["Custo total estimado"]
    df["Cancelado"] = df["Situação"].fillna("").str.casefold().eq("cancelado")
    return df


def filtrar(df: pd.DataFrame, inicio=None, fim=None, situacoes=None, canais=None, estados=None, produtos=None) -> pd.DataFrame:
    mascara = pd.Series(True, index=df.index)
    if inicio is not None and fim is not None:
        mascara &= df["Data"].dt.date.between(inicio, fim)
    if canais is not None:
        mascara &= df["Loja virtual"].isin(list(canais))
    if estados is not None:
        mascara &= df["Estado"].isin(list(estados))
    return df[mascara].copy()


def calcular_metricas_mes_atual_anterior(df: pd.DataFrame, canais=None, estados=None):
    df_f = filtrar(df, None, None, None, canais, estados, None)
    if df_f.empty or df_f["Data"].dropna().empty:
        return 0.0, 0.0, 0.0, 0.0, 0.0

    data_maxima = df_f["Data"].max()
    ano_atual = data_maxima.year
    mes_atual = data_maxima.month

    mes_ant_dt = data_maxima - pd.DateOffset(months=1)
    ano_ant = mes_ant_dt.year
    mes_ant = mes_ant_dt.month

    df_atual = df_f[(df_f["Data"].dt.year == ano_atual) & (df_f["Data"].dt.month == mes_atual)]
    df_ant = df_f[(df_f["Data"].dt.year == ano_ant) & (df_f["Data"].dt.month == mes_ant)]

    fat_liq_atual = float(df_atual["Preço Total"].sum())
    fat_liq_ant = float(df_ant["Preço Total"].sum())
    fat_bruto_atual = float(df_atual["Receita Produtos"].sum())
    fat_bruto_ant = float(df_ant["Receita Produtos"].sum())

    variacao = (fat_liq_atual - fat_liq_ant) / fat_liq_ant if fat_liq_ant > 0 else 0.0
    return fat_liq_atual, fat_liq_ant, fat_bruto_atual, fat_bruto_ant, variacao


def nivel_pedido(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["N° do Pedido", "Data", "Canal", "Situação", "Valor", "Frete"])
    return df.groupby("N° do Pedido", as_index=False).agg(
        Data=("Data", "min"), Canal=("Loja virtual", "first"), Situação=("Situação", "first"),
        Valor=("Preço Total", "sum"), Frete=("Frete", "first"),
    )


def indicadores(df: pd.DataFrame) -> dict:
    faturamento = float(df["Preço Total"].fillna(0).sum())
    pedidos = int(df["N° do Pedido"].nunique())
    return {
        "faturamento": faturamento,
        "pedidos": pedidos,
        "ticket_medio": faturamento / pedidos if pedidos else 0.0,
    }


def serie_diaria(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["Data", "Faturamento", "Pedidos"])
    d = df.groupby(df["Data"].dt.date).agg(Faturamento=("Preço Total", "sum"), Pedidos=("N° do Pedido", "nunique"))
    d.index = pd.to_datetime(d.index)
    d = d.reindex(pd.date_range(d.index.min(), d.index.max(), freq="D"), fill_value=0)
    d.index.name = "Data"
    return d.reset_index()


def curva_abc(df: pd.DataFrame) -> pd.DataFrame:
    p = df.groupby("Descrição", as_index=False).agg(Faturamento=("Preço Total", "sum"))
    p = p[p["Faturamento"] > 0].sort_values("Faturamento", ascending=False).reset_index(drop=True)
    total = p["Faturamento"].sum()
    if p.empty or total == 0:
        return p
    p["% do total"] = p["Faturamento"] / total * 100
    p["% acumulado"] = p["% do total"].cumsum()
    return p