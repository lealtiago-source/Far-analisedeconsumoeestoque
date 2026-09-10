"""Regressões dos módulos analise_remume / analise_orcamento / consumo."""
import datetime as dt
import unittest

import pandas as pd

from consumo import calcular_consumo
from planilhas import converter_datas


class TestConsumo(unittest.TestCase):
    def _df(self, linhas):
        df = pd.DataFrame(linhas, columns=["Medicamento/Produto", "Lote", "Data", "Quantidade"])
        df["Data"] = pd.to_datetime(df["Data"])
        df["Medicamento Agrupado"] = df["Medicamento/Produto"]
        return df

    def test_periodo_por_medicamento(self):
        """Item recente não deve ser diluído pelo período do item antigo."""
        df = self._df([
            ["ANTIGO", "L1", "2026-01-01", 100],
            ["ANTIGO", "L1", "2026-06-30", 100],
            ["NOVO", "L1", "2026-06-01", 300],
            ["NOVO", "L1", "2026-06-30", 300],
        ])
        r = calcular_consumo(df).set_index("Medicamento Agrupado")
        self.assertGreater(
            r.loc["NOVO", "Consumo Mensal Médio Corrigido"],
            r.loc["ANTIGO", "Consumo Mensal Médio Corrigido"],
        )

    def test_amostra_unica_nao_infla_30x(self):
        df = self._df([["X", "L1", "2026-03-10", 300]])
        r = calcular_consumo(df).iloc[0]
        self.assertEqual(r["Consumo Mensal Médio Corrigido"], 300)
        self.assertEqual(r["Amostra Insuficiente"], "SIM")

    def test_falta_nao_derruba_divisor_para_um_dia(self):
        df = self._df([
            ["X", "L1", "2026-01-01", 100],
            ["X", "L2", "2026-12-01", 100],
        ])
        r = calcular_consumo(df).iloc[0]
        self.assertGreaterEqual(r["Período Analisado (dias)"], int(335 * 0.25))

    def test_valor_unitario_ponderado(self):
        df = self._df([
            ["X", "L1", "2026-01-01", 10],
            ["X", "L1", "2026-03-01", 990],
        ])
        df["Valor Unitário"] = [100.0, 1.0]
        r = calcular_consumo(df, coluna_valor="Valor Unitário").iloc[0]
        self.assertLess(r["Valor Unitário Médio (R$)"], 10)


class TestDatas(unittest.TestCase):
    def test_iso_texto_e_datetime_misturados(self):
        s = pd.Series(["2026-01-16", dt.datetime(2026, 1, 31), "2026-02-15"])
        self.assertEqual(converter_datas(s).isna().sum(), 0)

    def test_formato_brasileiro(self):
        s = pd.Series(["31/01/2026", "15/02/2026"])
        convertidas = converter_datas(s)
        self.assertEqual(convertidas.isna().sum(), 0)
        self.assertEqual(convertidas.iloc[0].day, 31)


if __name__ == "__main__":
    unittest.main(verbosity=2)
