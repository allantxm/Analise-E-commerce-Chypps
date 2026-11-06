"""Testes da camada de dados. Na pasta do projeto:

    python -m unittest discover -s tests -v

Os testes com a base real só rodam se o fVendas.xlsx estiver ao lado do dashboard_chypps.py
(ou se a variável de ambiente BASE_XLSX apontar para ele). A base é aberta somente para
leitura e os testes conferem que o arquivo não foi alterado.
"""
import hashlib
import io
import os
import sys
import unittest
from pathlib import Path

import numpy as np
import pandas as pd

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ))
import dados  # noqa: E402


def bruto_sintetico() -> pd.DataFrame:
    """Imita a exportação: 3 canais, ID do Mercado Livre sem precisão, texto 'null' e datas em texto."""
    linhas = [
        # WooCommerce: pedido 100 com 2 linhas (total repetido), UF por extenso
        dict(n=100, estado="Concluído", data="2026-09-02 10:00:00", conta="WooCommerce", uf="São Paulo",
             bruto=1000.0, total=1100.0, un=1, sku="A1", titulo="Servidor A", custo=400.0, comprador="Ana Júlia Lopes",
             frete=100.0, nfe=None),
        dict(n=100, estado="Concluído", data="2026-09-02 10:00:00", conta="WooCommerce", uf="São Paulo",
             bruto=1000.0, total=1100.0, un=1, sku=None, titulo="Garantia", custo=None, comprador="Ana Júlia Lopes",
             frete=100.0, nfe=None),
        # Mercado Livre: mesmo N.º de venda (precisão perdida) em pedidos diferentes
        dict(n=2000020000000000, estado="Concluído", data="2026-09-05 09:00:00", conta="ML Comércio", uf="sp",
             bruto=500.0, total=450.0, un=1, sku="B1", titulo="Switch B", custo=None, comprador="Bruno",
             frete=None, nfe="Não emitida"),
        dict(n=2000020000000000, estado="Cancelada pelo comprador", data="2026-09-06 11:00:00", conta="ML Comércio",
             uf="RJ", bruto=300.0, total=0.0, un=1, sku="B2", titulo="Switch C", custo=None, comprador="Carla",
             frete=None, nfe="Não emitida"),
        dict(n=2000020000000000, estado="Concluído", data="2026-10-07 12:00:00", conta="ML Comércio", uf="MG",
             bruto=800.0, total=700.0, un=2, sku="B1", titulo="Switch B", custo=None, comprador="Dani",
             frete=None, nfe="Não emitida"),
        dict(n=2000020000000000, estado="Concluído", data="2026-10-07 12:00:00", conta="ML Comércio", uf="MG",
             bruto=800.0, total=700.0, un=2, sku="B1", titulo="Switch B", custo=None, comprador="Dani",
             frete=None, nfe="Não emitida"),
        dict(n=2000020000000000, estado="Enviado", data="2026-10-09 08:00:00", conta="ML Comércio", uf="PR",
             bruto=200.0, total=190.0, un=1, sku="B3", titulo="Cabo", custo=None, comprador="Edu",
             frete=None, nfe="Não emitida"),
        # Shopee: ID texto, data em texto dd/mm/aaaa, status 'null'
        dict(n="2609AAA", estado="Concluído", data="05/10/2026 10:30:00", conta="Shopee", uf="Bahia",
             bruto=100.0, total=90.0, un=1, sku="C1", titulo="Cabo", custo=None, comprador="Fábio",
             frete=0.0, nfe=None),
        dict(n="2609BBB", estado="null", data="15/10/2026 10:30:00", conta="Shopee", uf="null",
             bruto=50.0, total=45.0, un=1, sku="C1", titulo="Cabo", custo=None, comprador="Gil",
             frete=0.0, nfe=None),
    ]
    return pd.DataFrame({
        "N.º de venda": [x["n"] for x in linhas],
        "Estado": [x["estado"] for x in linhas],
        "Data de venda": [x["data"] for x in linhas],
        "Conta": [x["conta"] for x in linhas],
        "UF": [x["uf"] for x in linhas],
        "Cidade": ["Cidade"] * len(linhas),
        "Receita por produtos (BRL)": [x["bruto"] for x in linhas],
        "Total (BRL)": [x["total"] for x in linhas],
        "Unidades": [x["un"] for x in linhas],
        "SKU": [x["sku"] for x in linhas],
        "Título do anúncio": [x["titulo"] for x in linhas],
        "Custo do item": [x["custo"] for x in linhas],
        "Comprador": [x["comprador"] for x in linhas],
        "Receita por envio (BRL)": [x["frete"] for x in linhas],
        "NF-e em anexo": [x["nfe"] for x in linhas],
    })


class TestFormatacao(unittest.TestCase):
    def test_numero_br(self):
        r = dados.numero_br(pd.Series(["1.234,56", "R$ 94,9", "7", "x"], dtype="string"))
        self.assertAlmostEqual(r[0], 1234.56)
        self.assertAlmostEqual(r[1], 94.9)
        self.assertAlmostEqual(r[2], 7)
        self.assertTrue(np.isnan(r[3]))

    def test_formatos(self):
        self.assertEqual(dados.brl(1234567.891), "R$ 1.234.567,89")
        self.assertEqual(dados.inteiro(20341), "20.341")
        self.assertEqual(dados.pct(0.0705), "7,0%")

    def test_rotulo_mom_como_no_dax(self):
        self.assertEqual(dados.rotulo_mom(0.4918), "📈 +49,18%")
        self.assertEqual(dados.rotulo_mom(-0.05), "📉 −5,00%")
        self.assertEqual(dados.rotulo_mom(0), "⚪ 0,00%")
        self.assertEqual(dados.rotulo_mom(float("nan")), "⚪ 0,00%")

    def test_mascara_nome(self):
        self.assertEqual(dados.mascarar_nome("Ana Júlia Lopes"), "A. J. L.")
        self.assertEqual(dados.mascarar_nome(None), "")
        self.assertEqual(dados.mascarar_nome(float("nan")), "")

    def test_uf(self):
        s = dados.normalizar_uf(pd.Series(["São Paulo", "sao paulo", "SP", "null", None, "Paraná"]))
        self.assertEqual(list(s), ["SP", "SP", "SP", dados.NAO_INFORMADO, dados.NAO_INFORMADO, "PR"])

    def test_grupos_de_situacao(self):
        g = dados.grupo_situacao
        self.assertEqual(g("Concluído"), "Concluído")
        self.assertEqual(g("Cancelada pelo comprador"), "Cancelado")
        self.assertEqual(g("Malsucedido"), "Cancelado")
        self.assertEqual(g("Você cancelou a venda"), "Cancelado")
        self.assertEqual(g("Devolução finalizada com reembolso para o comprador"), "Devolvido/reembolsado")
        self.assertEqual(g("Reembolsado"), "Devolvido/reembolsado")
        self.assertEqual(g("Devolução a caminho"), "Devolvido/reembolsado")
        # prazo de devolução em aberto e mediação a favor do vendedor não são devolução
        self.assertEqual(g("O comprador pode pedir uma devolução até 2026-10-05"), "Concluído")
        self.assertEqual(g("Mediação finalizada. Te demos o dinheiro."), "Concluído")
        self.assertEqual(g("A caminho"), "Em andamento")
        self.assertEqual(g(None), "Em andamento")


class TestSintetico(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.df = dados.preparar(bruto_sintetico())

    def test_leitura_e_tipos(self):
        d = self.df
        self.assertEqual(len(d), 9)
        self.assertEqual(d["Data"].isna().sum(), 0)
        self.assertEqual(d.loc[d["Loja virtual"] == "Shopee", "Data"].min(), pd.Timestamp("2026-10-05"))  # dia primeiro
        self.assertFalse(d["Descrição"].astype(str).str.contains("nan|null", case=False).any())
        self.assertEqual(d["Situação"].iloc[-1], "Sem situação")        # 'null' não vira texto
        self.assertEqual(d["Estado"].iloc[0], "SP")
        self.assertEqual(d["Estado"].iloc[-1], dados.NAO_INFORMADO)
        self.assertEqual(d["Código do produto"].iloc[1], "(sem SKU)")
        self.assertEqual(d["Situação da NFe"].iloc[2], "Pendente")

    def test_pedidos_do_mercado_livre_sao_reconstruidos(self):
        ml = self.df[self.df["Loja virtual"] == "ML Comércio"]
        # 5 linhas, 4 pedidos (as duas linhas idênticas de 07/10 são o mesmo pedido)
        self.assertEqual(len(ml), 5)
        self.assertEqual(ml["N° do Pedido"].nunique(), 4)
        self.assertEqual(self.df["N° do Pedido"].nunique(), 4 + 1 + 2)

    def test_bruto_e_liquido_como_no_dax(self):
        self.assertAlmostEqual(dados.faturamento_bruto(self.df), 1000 * 2 + 500 + 300 + 800 * 2 + 200 + 100 + 50)
        self.assertAlmostEqual(dados.faturamento_liquido(self.df), 1100 * 2 + 450 + 0 + 700 * 2 + 190 + 90 + 45)
        ind = dados.indicadores(self.df)
        self.assertAlmostEqual(ind["faturamento_liquido"], dados.faturamento_liquido(self.df))
        self.assertAlmostEqual(ind["faturamento"], ind["faturamento_liquido"])

    def test_mom_mes_atual_e_anterior(self):
        m = dados.metricas_mom(self.df)
        self.assertEqual((m["mes_atual"], m["mes_anterior"]), ("2026-10", "2026-09"))
        self.assertAlmostEqual(m["liquido_atual"], 700 * 2 + 190 + 90 + 45)
        self.assertAlmostEqual(m["liquido_anterior"], 1100 * 2 + 450)
        self.assertAlmostEqual(m["var_liquido"], (m["liquido_atual"] - m["liquido_anterior"]) / m["liquido_anterior"])
        self.assertTrue(m["rotulo_liquido"].startswith("📉"))
        self.assertAlmostEqual(m["bruto_atual"], 800 * 2 + 200 + 100 + 50)

    def test_mom_so_respeita_canal_e_situacao(self):
        # período/UF/produto não entram (ALLEXCEPT): filtrar por UF não muda o MoM
        so_ml = dados.metricas_mom(self.df, canais=["ML Comércio"])
        self.assertAlmostEqual(so_ml["liquido_anterior"], 450)
        self.assertAlmostEqual(so_ml["liquido_atual"], 700 * 2 + 190)
        sem_cancel = dados.metricas_mom(self.df, situacoes=["Concluído"])
        self.assertAlmostEqual(sem_cancel["liquido_anterior"], 1100 * 2 + 450)
        # canal sem mês anterior: DIVIDE(..., 0) devolve 0 => ⚪
        so_shopee = dados.metricas_mom(self.df, canais=["Shopee"])
        self.assertFalse(so_shopee["tem_anterior"])
        self.assertEqual(so_shopee["var_liquido"], 0.0)
        self.assertEqual(so_shopee["rotulo_liquido"], "⚪ 0,00%")
        vazio = dados.metricas_mom(self.df, canais=[])
        self.assertEqual(vazio["liquido_atual"], 0.0)

    def test_compatibilidade_funcao_antiga(self):
        liq_a, liq_p, bru_a, bru_p, var = dados.calcular_metricas_mes_atual_anterior(self.df, canais=["ML Comércio"])
        self.assertAlmostEqual(liq_p, 450)
        self.assertAlmostEqual(var, (liq_a - liq_p) / liq_p)

    def test_resumo_mensal(self):
        r = dados.resumo_mensal(self.df)
        self.assertEqual(list(r["Mês"]), ["2026-09", "2026-10"])
        self.assertTrue(np.isnan(r["MoM líquido"].iloc[0]))
        self.assertAlmostEqual(r["MoM líquido"].iloc[1], dados.metricas_mom(self.df)["var_liquido"])

    def test_frete_por_pedido_e_pct_zero(self):
        ind = dados.indicadores(self.df)
        self.assertAlmostEqual(ind["frete_total"], 100.0)      # frete do pedido 100 conta uma vez
        self.assertAlmostEqual(ind["pct_frete_gratis"], 2 / 3)  # só entre os pedidos com frete informado

    def test_margem_so_onde_ha_custo(self):
        ind = dados.indicadores(self.df)
        self.assertAlmostEqual(ind["margem_bruta"], 1000 - 400)
        self.assertAlmostEqual(ind["margem_pct"], 0.6)
        self.assertLess(ind["cobertura_custo"], 1.0)
        p = dados.desempenho_produtos(self.df).set_index("Descrição")
        self.assertTrue(np.isnan(p.loc["Switch B", "Margem"]))
        self.assertFalse(np.isnan(p.loc["Servidor A", "Margem"]))

    def test_cancelamento_e_devolucao(self):
        self.assertEqual(int(self.df["Cancelado"].sum()), 1)
        t = dados.taxa_cancelamento(self.df, "Loja virtual").set_index("Loja virtual")
        self.assertAlmostEqual(t.loc["ML Comércio", "Taxa de cancelamento"], 1 / 4)
        self.assertAlmostEqual(dados.taxa_cancelamento(self.df), 1 / 7)

    def test_filtros(self):
        d = self.df
        self.assertEqual(len(dados.filtrar(d, situacoes=["Cancelado"])), 1)
        self.assertEqual(len(dados.filtrar(d, canais=["Shopee"])), 2)
        self.assertEqual(len(dados.filtrar(d, estados=["MG"])), 2)
        self.assertEqual(len(dados.filtrar(d, produtos=["Cabo"])), 3)
        self.assertEqual(len(dados.filtrar(d, produtos=[])), 9)      # lista vazia = todos
        dia = pd.Timestamp("2026-09-02").date()
        self.assertEqual(len(dados.filtrar(d, inicio=dia, fim=dia)), 2)

    def test_serie_diaria(self):
        s = dados.serie_diaria(self.df)
        self.assertEqual(len(s), 44)   # 02/09 a 15/10
        self.assertAlmostEqual(s["Acumulado"].iloc[-1], dados.faturamento_liquido(self.df))
        self.assertIn("Média móvel 7 dias", s)

    def test_curva_abc_e_destaques(self):
        abc = dados.curva_abc(self.df)
        self.assertAlmostEqual(abc["% acumulado"].iloc[-1], 100.0)
        self.assertEqual(abc["Classe"].iloc[0], "A")
        self.assertTrue(dados.destaques(self.df, self.df))

    def test_qualidade(self):
        q = dados.qualidade(self.df)
        self.assertEqual(q["linhas_identicas"], 1)       # só as duas linhas iguais do ML em 07/10
        self.assertEqual(q["pedidos_total_repetido"], 2)
        self.assertEqual(q["uf_nao_informada"], 1)
        self.assertEqual(len(self.df), 9)                 # nada é removido

    def test_vazio_nao_quebra(self):
        vazio = dados.filtrar(self.df, canais=["Inexistente"])
        self.assertTrue(vazio.empty)
        self.assertEqual(dados.indicadores(vazio)["pedidos"], 0)
        self.assertEqual(dados.destaques(vazio), [])
        self.assertTrue(dados.serie_diaria(vazio).empty)
        self.assertTrue(dados.curva_abc(vazio).empty)
        self.assertTrue(dados.resumo_mensal(vazio).empty)
        self.assertTrue(dados.desempenho_produtos(vazio).empty)
        self.assertTrue(dados.nivel_pedido(vazio).empty)

    def test_xlsx_em_bytes(self):
        buf = io.BytesIO()
        bruto = bruto_sintetico().replace("null", np.nan)
        bruto.to_excel(buf, index=False)
        d = dados.carregar(buf.getvalue(), "xlsx")
        self.assertEqual(len(d), 9)

    def test_coluna_obrigatoria_ausente(self):
        with self.assertRaises(ValueError):
            dados.preparar(bruto_sintetico().drop(columns=["Total (BRL)"]))


def _achar_xlsx():
    for c in (os.environ.get("BASE_XLSX"), RAIZ / "fVendas.xlsx"):
        if c and Path(c).exists():
            return Path(c)
    return None


@unittest.skipIf(_achar_xlsx() is None, "fVendas.xlsx não encontrado ao lado do dashboard_chypps.py")
class TestBaseReal(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.caminho = _achar_xlsx()
        cls.hash_antes = hashlib.sha256(cls.caminho.read_bytes()).hexdigest()
        cls.df = dados.carregar(cls.caminho.read_bytes())

    def test_totais_batem_com_o_power_bi(self):
        self.assertEqual(len(self.df), 296)
        self.assertAlmostEqual(dados.faturamento_bruto(self.df), 626056.24, places=2)
        self.assertAlmostEqual(dados.faturamento_liquido(self.df), 518846.98, places=2)

    def test_mes_atual_e_anterior(self):
        m = dados.metricas_mom(self.df)
        self.assertEqual((m["mes_atual"], m["mes_anterior"]), ("2026-10", "2026-09"))
        self.assertAlmostEqual(m["bruto_atual"], 368802.52, places=2)
        self.assertAlmostEqual(m["bruto_anterior"], 257253.72, places=2)
        self.assertAlmostEqual(m["liquido_atual"], 310628.03, places=2)
        self.assertAlmostEqual(m["liquido_anterior"], 208218.95, places=2)
        self.assertEqual(m["rotulo_liquido"], "📈 +49,18%")
        self.assertEqual(m["rotulo_bruto"], "📈 +43,36%")

    def test_pedidos_e_canais(self):
        self.assertEqual(self.df["N° do Pedido"].nunique(), 290)
        por_canal = self.df.groupby("Loja virtual")["N° do Pedido"].nunique().to_dict()
        self.assertEqual(por_canal, {"ML BJS": 13, "ML Comércio": 90, "ML Serviços": 80,
                                     "Shopee": 67, "WooCommerce": 40})

    def test_limpeza(self):
        self.assertEqual(self.df["Data"].isna().sum(), 0)
        self.assertEqual(self.df["Data"].min(), pd.Timestamp("2026-09-01"))
        self.assertEqual(self.df["Data"].max(), pd.Timestamp("2026-10-31"))
        self.assertTrue(set(self.df["Estado"]) <= set(dados.CENTROS_UF) | {dados.NAO_INFORMADO})
        for col in ("Descrição", "Município", "Código do produto", "Situação"):
            self.assertFalse(self.df[col].astype(str).str.casefold().isin(["nan", "null", "none"]).any(), col)

    def test_base_nao_foi_alterada(self):
        dados.filtrar(self.df, canais=["Shopee"])
        dados.indicadores(self.df)
        self.assertEqual(hashlib.sha256(self.caminho.read_bytes()).hexdigest(), self.hash_antes)


if __name__ == "__main__":
    unittest.main()
