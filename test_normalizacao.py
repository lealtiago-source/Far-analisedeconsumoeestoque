"""
Testes de regressão da normalização de medicamentos.

Cada teste documenta um bug real encontrado na versão anterior.
Rodar com: python -m unittest discover -v
"""

import unittest

from normalizacao import agrupar_medicamento as norm


class TestNormalizacao(unittest.TestCase):


    def test_remove_sal_e_forma_farmaceutica(self):
        assert norm("LOSARTANA POTÁSSICA 50 MG COMPRIMIDO REVESTIDO") == "LOSARTANA 50 MG"


    def test_padroniza_concentracao_sem_espaco(self):
        assert norm("DIPIRONA 500MG") == "DIPIRONA 500 MG"
        assert norm("METFORMINA 850MG COMP REVESTIDO") == "METFORMINA 850 MG"


    def test_diluicao_nao_e_apagada(self):
        """
        Bug anterior: '5 ML' estava na lista de apresentações removidas, então
        5 MG/5 ML e 5 MG/1 ML colapsavam na mesma chave.
        """
        assert norm("MIDAZOLAM 5MG/5ML SOLUCAO INJETAVEL") != norm(
            "MIDAZOLAM 5MG/1ML SOLUCAO INJETAVEL"
        )


    def test_principio_ativo_com_nome_de_sal_e_preservado(self):
        """Bug anterior: 'CARBONATO' era removido e sobrava 'DE LITIO'."""
        assert norm("CARBONATO DE LITIO 300 MG COMPRIMIDO") == "CARBONATO DE LITIO 300 MG"
        assert norm("SULFATO DE ZINCO 4MG/ML") == "SULFATO DE ZINCO 4 MG/ML"


    def test_corrige_mojibake_e_nbsp(self):
        assert norm("LosartanaÂ 50 MG") == "LOSARTANA 50 MG"
        assert norm("SoluÃ§Ã£o dipirona 500 MG") == "DIPIRONA 500 MG"


    def test_remove_numeracao_inicial(self):
        assert norm("1 - DIPIRONA 500 MG") == norm("DIPIRONA 500 MG")
        assert norm("12) DIPIRONA 500 MG") == norm("DIPIRONA 500 MG")


    def test_preserva_marcadores(self):
        assert "FR" in norm("DIPIRONA 500 MG/ML (FR)")
        assert "RF/CANETA" in norm("INSULINA 100 UI/ML (RF/CANETA)")


    def test_idempotencia(self):
        """A chave normalizada deve ser estável ao ser normalizada novamente."""
        entradas = [
            "LOSARTANA POTÁSSICA 50 MG COMPRIMIDO REVESTIDO",
            "1 - Dipirona sódica 500mg/ml solução injetável (FR)",
            "MIDAZOLAM 5MG/5ML SOLUCAO INJETAVEL",
            "CARBONATO DE LITIO 300 MG COMPRIMIDO",
        ]
        for entrada in entradas:
            chave = norm(entrada)
            assert norm(chave) == chave, entrada


    def test_valores_nulos(self):
        assert norm(None) == ""
        assert norm("") == ""
        assert norm("   ") == ""


if __name__ == "__main__":
    unittest.main(verbosity=2)
