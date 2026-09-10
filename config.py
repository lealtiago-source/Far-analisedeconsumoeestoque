"""
Configuração central da aplicação.

Antes: caminhos ("static", "uploads") escritos como strings relativas dentro
de cada módulo de análise. Isso quebrava quando o processo rodava com um
diretório de trabalho diferente — o caso do executável empacotado (.exe),
em que os arquivos eram gravados na pasta errada e o download retornava 404.

Agora todo caminho vem daqui e é sempre absoluto.
"""

from __future__ import annotations

import os
import sys
from pathlib import Path


def _base_dir() -> Path:
    """Pasta ao lado do executável (.exe) ou raiz do projeto (python)."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def _internal_dir() -> Path:
    """Pasta de recursos embutidos (templates) — só existe no .exe."""
    if getattr(sys, "frozen", False):
        return Path(getattr(sys, "_MEIPASS", _base_dir()))
    return Path(__file__).resolve().parent


BASE_DIR = _base_dir()
INTERNAL_DIR = _internal_dir()

TEMPLATE_DIR = INTERNAL_DIR / "templates"
STATIC_DIR = BASE_DIR / "static"
UPLOAD_DIR = BASE_DIR / "uploads"
# Resultados ficam FORA de static/: são servidos por rota autenticada.
RESULTADO_DIR = BASE_DIR / "resultados"
LOG_DIR = BASE_DIR / "logs"

for _pasta in (STATIC_DIR, UPLOAD_DIR, RESULTADO_DIR, LOG_DIR):
    _pasta.mkdir(parents=True, exist_ok=True)

# --- Parâmetros de negócio (antes hardcoded no meio da lógica) -------------

# Meses de cobertura usados no cálculo da quantidade a comprar.
MESES_COBERTURA = int(os.environ.get("MESES_COBERTURA", 8))

# Janela em dias para considerar validade crítica / fim de estoque.
DIAS_VALIDADE_PADRAO = int(os.environ.get("DIAS_VALIDADE_PADRAO", 120))
DIAS_ESTOQUE_PADRAO = int(os.environ.get("DIAS_ESTOQUE_PADRAO", 120))

# Intervalo (dias) sem movimentação que passa a contar como desabastecimento.
DIAS_MINIMOS_DESABASTECIMENTO = int(os.environ.get("DIAS_MINIMOS_DESABASTECIMENTO", 15))

# Janela mínima de observação. Um medicamento com uma única movimentação tem
# período de 1 dia; projetar isso para 30 dias multiplicava o consumo por 30 e
# gerava quantidades de compra irreais. Períodos menores que este valor são
# tratados como este valor.
DIAS_MINIMOS_PERIODO = int(os.environ.get("DIAS_MINIMOS_PERIODO", 30))

# Teto de segurança: proporção mínima do período que deve permanecer válida
# após descontar dias de falta. Evita dividir o consumo por pouquíssimos dias
# e inflar o consumo mensal (e a compra) em várias vezes.
PROPORCAO_MINIMA_PERIODO_VALIDO = 0.25

MAX_UPLOAD_BYTES = int(os.environ.get("MAX_UPLOAD_BYTES", 50 * 1024 * 1024))
EXTENSOES_PERMITIDAS = {".xlsx", ".xls"}
