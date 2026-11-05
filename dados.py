"""Camada de dados do dashboard Chypps Tecnologia."""
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
    bruto = pd.read_excel(io.BytesIO(conteudo), dtype=str)
    bruto.columns = bruto.columns.astype(str).str.strip()
    return adaptar_exportacao_original(bruto)


def ler_csv(conteudo: bytes) -> pd.DataFrame:
    for codificacao in ("utf-8-sig", "latin-1"):
        for sep in (";", ","):
            try:
                bruto = pd.read_csv(io.BytesIO(conteudo), sep=sep, dtype="string", encoding=codificacao)
                bruto.columns = bruto.columns.astype(str).str.strip()
                return adaptar_exportacao_original(bruto)
            except Exception:
                continue
    raise ValueError("Não foi possível interpretar o arquivo.")


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
        mascara &= df["Data"].dt.date.between(inicio, fim)
    if canais is not None:
        mascara &= df["Loja virtual"].isin(list(canais))
    if estados is not None:
        mascara &= df["Estado"].isin(list(estados))
    return df[mascara].copy()


def indicadores(df: pd.DataFrame) -> dict:
    faturamento_liquido = float(df["Preço Total"].fillna(0).sum())
    faturamento_bruto = float(df["Faturamento Bruto"].fillna(0).sum()) if "Faturamento Bruto" in df.columns else faturamento_liquido
    pedidos = int(df["N° do Pedido"].nunique())
    
    return {
        "faturamento": faturamento_liquido,
        "faturamento_bruto": faturamento_bruto,
        "pedidos": pedidos,
        "ticket_medio": faturamento_liquido / pedidos if pedidos else 0.0,
    }