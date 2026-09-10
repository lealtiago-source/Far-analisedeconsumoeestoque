"""
Leitura defensiva de planilhas.

Substitui `ler_xml_robusto` e `converter_quantidade` de analise.py.

Problemas corrigidos:
  - `df_test.iloc[1]` estourava IndexError em planilhas com uma única linha.
  - `converter_quantidade` truncava frações (int(float(x))) e engolia
    qualquer erro retornando 0 — silenciosamente perdia quantidades.
  - "1.234,56" virava "1.234.56" e caía no zero silencioso.
  - Nenhuma validação de colunas obrigatórias: uma coluna renomeada na
    origem produzia KeyError capturado por um `except` genérico, e o
    usuário recebia "Análise concluída" sem análise nenhuma.
"""

from __future__ import annotations

import logging
import re
from pathlib import Path

import pandas as pd

logger = logging.getLogger(__name__)


class PlanilhaInvalidaError(ValueError):
    """Erro de negócio: a planilha não tem o formato esperado."""


_NUMERO_RE = re.compile(r"[^\d,.\-]")


def ler_planilha(caminho: str | Path, sheet_name: object = 0, **kwargs) -> pd.DataFrame:
    """
    Lê uma planilha detectando automaticamente se o cabeçalho está na
    primeira ou na segunda linha (relatórios do SIGAF variam).
    """
    caminho = Path(caminho)
    if not caminho.exists():
        raise PlanilhaInvalidaError(f"Arquivo não encontrado: {caminho.name}")

    bruto = pd.read_excel(caminho, sheet_name=sheet_name, header=None, **kwargs)

    header = 0
    # Só compara se houver ao menos duas linhas (antes: IndexError).
    if len(bruto) >= 2 and bruto.iloc[0].count() < bruto.iloc[1].count():
        header = 1

    df = pd.read_excel(caminho, sheet_name=sheet_name, header=header, **kwargs)
    df.columns = [str(c).strip() for c in df.columns]
    return df.dropna(how="all").reset_index(drop=True)


def exigir_colunas(df: pd.DataFrame, colunas: list[str], origem: str) -> None:
    """Falha cedo e com mensagem clara quando a planilha muda de formato."""
    faltando = [c for c in colunas if c not in df.columns]
    if faltando:
        raise PlanilhaInvalidaError(
            f"A planilha de {origem} não contém a(s) coluna(s): "
            f"{', '.join(faltando)}. Colunas encontradas: {', '.join(df.columns)}"
        )


def converter_numero(valor: object) -> float:
    """
    Converte um valor de planilha em float, aceitando formato brasileiro.

    Diferente da versão anterior: não trunca frações e registra em log os
    valores que não puderam ser convertidos, em vez de virarem zero em silêncio.
    """
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return 0.0
    if isinstance(valor, bool):
        return float(int(valor))
    if isinstance(valor, (int, float)):
        return float(valor)

    texto = _NUMERO_RE.sub("", str(valor).strip())
    if not texto:
        return 0.0

    # "1.234,56" -> "1234.56" | "1234,56" -> "1234.56" | "1234.56" mantém
    if "," in texto:
        texto = texto.replace(".", "").replace(",", ".")

    try:
        return float(texto)
    except ValueError:
        logger.warning("Valor numérico ignorado (não convertível): %r", valor)
        return 0.0


def converter_quantidade(valor: object) -> int:
    """Quantidade em unidades — arredonda, não trunca."""
    return int(round(converter_numero(valor)))


def converter_datas(serie: pd.Series, origem: str = "") -> pd.Series:
    """
    Converte uma coluna de datas aceitando ISO e formato brasileiro na MESMA
    coluna, sem perder linhas em silêncio.

    Por que existe: `pd.to_datetime(col, dayfirst=True)` sobre uma coluna que
    mistura texto ISO ("2026-01-16") com datetime do Excel devolve NaT para a
    maioria das linhas — e o cálculo seguia com uma fração dos dados, gerando
    período e consumo errados sem nenhum aviso.
    """
    bruto = pd.to_datetime(serie, errors="coerce", format="mixed", dayfirst=False)

    faltando = bruto.isna() & serie.notna()
    if faltando.any():
        # Segunda tentativa só nas linhas restantes, agora dia primeiro.
        retentativa = pd.to_datetime(serie[faltando], errors="coerce", dayfirst=True)
        bruto = bruto.copy()
        bruto[faltando] = retentativa

    perdidas = int((bruto.isna() & serie.notna()).sum())
    if perdidas:
        logger.warning(
            "%d data(s) não reconhecida(s)%s foram ignoradas.",
            perdidas,
            f" em {origem}" if origem else "",
        )
    return bruto
