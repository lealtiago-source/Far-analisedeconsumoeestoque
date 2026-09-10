"""
Normalização e agrupamento de nomes de medicamentos.

Substitui o antigo utils.py.

Principais mudanças:
  - Regex pré-compiladas no import (antes: ~100 recompilações por chamada).
  - Cache LRU: nomes repetidos custam O(1) (planilhas repetem muito).
  - Remoção de sais com proteção por lista branca (evita destruir
    princípios ativos como CARBONATO DE LITIO / SULFATO DE ZINCO).
  - Não remove mais volumes ("5 ML") que fazem parte da concentração
    (ex.: 5 MG/5 ML deixou de ser confundido com 5 MG/1 ML).
  - Função `agrupar_medicamento` definida UMA vez (antes: duplicada).
"""

from __future__ import annotations

import re
import unicodedata
from functools import lru_cache

# ---------------------------------------------------------------------------
# Correção de mojibake (texto UTF-8 lido como Latin-1)
# ---------------------------------------------------------------------------

_MOJIBAKE = {
    "Ã¡": "á", "Ã ": "à", "Ã¢": "â", "Ã£": "ã", "Ã¤": "ä",
    "Ã©": "é", "Ãª": "ê", "Ã­": "í", "Ã³": "ó", "Ã´": "ô",
    "Ãµ": "õ", "Ãº": "ú", "Ã§": "ç",
    "Â ": " ", "Â": "",  # nbsp lido como Latin-1
}

_CONTROLE_RE = re.compile(r"[\x80-\x9f]")
_ESPACOS_RE = re.compile(r"\s+")

# ---------------------------------------------------------------------------
# Unidades de concentração
# ---------------------------------------------------------------------------

_UNIDADES = "MG|MCG|G|ML|UI|MEQ|MMOL"

# 500MG -> 500 MG   (\b garante que não quebre "5MGX")
_NUM_UNIDADE_RE = re.compile(rf"(\d)\s*({_UNIDADES})\b", re.IGNORECASE)
# normaliza os espaços ao redor da barra: "500 MG / ML" -> "500 MG/ML"
_BARRA_RE = re.compile(r"\s*/\s*")


def normalizar_texto(texto: object) -> str:
    """Remove acentos, mojibake, caracteres de controle e padroniza unidades."""
    if texto is None:
        return ""

    texto = str(texto)

    for antigo, novo in _MOJIBAKE.items():
        texto = texto.replace(antigo, novo)

    texto = _CONTROLE_RE.sub("", texto)

    # Remove acentos preservando a letra base.
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))

    texto = _ESPACOS_RE.sub(" ", texto)
    texto = _NUM_UNIDADE_RE.sub(r"\1 \2", texto)
    texto = _BARRA_RE.sub("/", texto)

    return texto.strip()


# ---------------------------------------------------------------------------
# Marcadores preservados durante a limpeza
# ---------------------------------------------------------------------------

_MARCADORES = {
    "(FR)": "\x00FR\x00",
    "(RF/CANETA)": "\x00RFCANETA\x00",
}

_RESTAURAR_MARCADORES = {
    "\x00FR\x00": " FR",
    "\x00RFCANETA\x00": " RF/CANETA",
}

# Princípios ativos em que o "sal" faz parte do nome: são protegidos por
# placeholder antes da remoção de sais e restaurados depois.
_PROTEGIDOS = [
    "SULFATO FERROSO",
    "SULFATO DE MAGNESIO",
    "SULFATO DE ZINCO",
    "SULFATO DE ATROPINA",
    "CLORETO DE POTASSIO",
    "CLORETO DE SODIO",
    "CLORETO DE CALCIO",
    "CARBONATO DE CALCIO",
    "CARBONATO DE LITIO",
    "ACETATO DE MEDROXIPROGESTERONA",
    "CITRATO DE SILDENAFILA",
    "NITRATO DE MICONAZOL",
    "FOSFATO DE OSELTAMIVIR",
    "PERMANGANATO DE POTASSIO",
]

_PROTEGIDOS_MAP = {
    nome: f"\x00P{i}\x00" for i, nome in enumerate(sorted(_PROTEGIDOS, key=len, reverse=True))
}
_PROTEGIDOS_INVERSO = {v: k for k, v in _PROTEGIDOS_MAP.items()}

# ---------------------------------------------------------------------------
# Listas de termos removidos / padronizados
# ---------------------------------------------------------------------------

_NUMERACAO_INICIAL_RE = re.compile(r"^\s*\d+\s*(?:[-.)])\s*")
_PARENTESES_RE = re.compile(r"\(.*?\)")

_ROTULOS = [
    "PRINCIPIO ATIVO", "FORMA FARMACEUTICA", "CONCENTRACAO", "APRESENTACAO",
    "COMPOSICAO ASSOCIADA COM", "INDICACAO", "APLICACAO", "COMPOSICAO", "DOSAGEM",
]

_SAIS = [
    "CLORIDRATO DE", "CLORIDRATO", "FOSFATO DISSODICO DE", "SUCCINATO DE SODIO",
    "BUTILBROMETO DE", "BUTILBROMETO", "FOSFATO DE", "SULFATO DE", "DIPROPIONATO",
    "DIPROPRIONATO", "HEMIFUMARATO", "HEMITARTARATO", "MONOIDRATADA", "MONONITRATO",
    "BISSULFATO", "DINITRATO", "POTASSICO", "POTASSICA", "DISSODICO", "CARBONATO",
    "SUCCINATO", "BESILATO", "MALEATO", "CLORETO", "OXALATO", "ACETATO", "NITRATO",
    "CITRATO", "FOSFATO", "SULFATO", "SODICO", "SODICA", "SAL",
]

_INJETAVEL = [
    "PO LIOFILIZADO PARA SOLUCAO INJETAVEL", "PO LIOFILICO PARA INJETAVEL",
    "PO PARA SUSPENSAO INJETAVEL", "PO PARASOLUCAO INJETAVEL",
    "PO PARA SOLUCAO INJETAVEL", "PO PARA INJETAVEL", "SUSPENSAO INJETAVEL",
    "SOLUCAO INJETAVEL", "ORAL INJETAVEL", "SOL INJETAVEL", "AMPOLAS",
    "AMPOLA", "SOL INJ", "INJETAVEL", "INJET", "AMP", "INJ",
]

_ORAL = [
    "PO PARA SUSPENSAO ORAL", "PO P SUSPENSAO ORAL", "SUSPENSAO ORAL",
    "SOLUCAO ORAL", "PO PARA ORAL", "EMULSAO ORAL", "SOL ORAL", "EMULSAO",
    "XAROPE", "XAROP", "FRASCOS", "FRASCO", "FRACO", "GOTAS", "GOTA",
]

_COMPRIMIDO = [
    "COMPRIMIDOS REVESTIDOS DE LIBERACAO PROLONGADA",
    "COMPRIMIDO REVESTIDO DE LIBERACAO PROLONGADA",
    "COMPRIMIDOS REVESTIDOS DE LIBERACAO RETARDADA",
    "COMPRIMIDO REVESTIDO DE LIBERACAO RETARDADA",
    "COMPRIMIDOS REVESTIDOS DE LIBERACAO CONTROLADA",
    "COMPRIMIDO REVESTIDO DE LIBERACAO CONTROLADA",
    "COMPRIMIDOS DE LIBERACAO PROLONGADA",
    "COMPRIMIDO DE LIBERACAO PROLONGADA",
    "DE LIBERACAO RETARDADA",
    "COMPRIMIDOS REVESTIDOS", "COMPRIMIDO REVESTIDO",
    "CAPSULAS DURAS", "CAPSULA DURA", "CAPSULAS MOLES", "CAPSULA MOLE",
    "COMPRIMIDOS", "COMPRIMIDO", "CAPSULAS", "CAPSULA",
    "DRAGEAS", "DRAGEA", "REVESTIDO", "COMP",
]

# ATENÇÃO: volumes ("5 ML") foram removidos desta lista de propósito —
# eles fazem parte da concentração (5 MG/5 ML) e apagá-los agrupava
# apresentações clinicamente diferentes.
_APRESENTACAO = [
    "BLISTERS", "BLISTER", "CAIXAS", "CAIXA", "UNIDADES", "UNIDADE",
    "PO SUSP O", "SUSPENSAO", "SUSP", "DERMATO", "SOLUCAO", "TIPO",
    "VIT B9", "USO",
]


def _compilar_grupo(termos: list[str]) -> re.Pattern[str]:
    """Compila uma alternância com \\b, do termo mais longo para o mais curto."""
    ordenados = sorted(set(termos), key=len, reverse=True)
    return re.compile(r"\b(?:" + "|".join(re.escape(t) for t in ordenados) + r")\b")


_ROTULOS_RE = _compilar_grupo(_ROTULOS)
_SAIS_RE = _compilar_grupo(_SAIS)
_INJETAVEL_RE = _compilar_grupo(_INJETAVEL)
_ORAL_RE = _compilar_grupo(_ORAL)
_COMPRIMIDO_RE = _compilar_grupo(_COMPRIMIDO)
_APRESENTACAO_RE = _compilar_grupo(_APRESENTACAO)

_PONTUACAO_RE = re.compile(r"[,;.:\-]")
_REPETICAO_RE = re.compile(r"\b(\w+)(\s+\1\b)+")

# Equivalências manuais aplicadas ao final (chave já normalizada).
_EQUIVALENCIAS = {
    "SAIS DE REIDRATACAO ORAL 1 066+6 093+0 884+0 457 G PO PARA PREPARACAO EXTEMPORANEA":
        "SAIS DE REIDRATACAO ORAL 3 5+20 0+2 9+1 5 G PO PARA PREPARACAO EXTEMPORANEA",
    "TIAMINA + PIRIDOXINA + CIANOCOBALAMINA 100+100+5 MG/ML INJETAVEL":
        "CIANOCOBALAMINA + PIRIDOXINA + TIAMINA 100+100+5 MG/ML INJETAVEL",
    "BETAMETASONA + BETAMETASONA 6 43+2 63 MG/ML INJETAVEL":
        "BETAMETASONA + BETAMETASONA 5+2 MG/ML INJETAVEL",
}


@lru_cache(maxsize=100_000)
def normalizar_medicamento(nome: object) -> str:
    """
    Reduz um nome de medicamento a uma chave canônica comparável entre
    planilhas de origens diferentes (dispensação, estoque, licitação).

    Idempotente: normalizar_medicamento(normalizar_medicamento(x)) == normalizar_medicamento(x).
    """
    if nome is None:
        return ""

    nome = normalizar_texto(nome).upper()
    if not nome:
        return ""

    for original, marcador in _MARCADORES.items():
        nome = nome.replace(original, marcador)

    nome = _NUMERACAO_INICIAL_RE.sub("", nome)
    nome = _PARENTESES_RE.sub(" ", nome)
    nome = nome.replace(":", " ")
    nome = nome.replace("ASSOCIADA COM", "+")
    nome = _ROTULOS_RE.sub(" ", nome)

    for original, marcador in _PROTEGIDOS_MAP.items():
        nome = nome.replace(original, marcador)

    nome = _SAIS_RE.sub(" ", nome)
    nome = _INJETAVEL_RE.sub(" INJETAVEL ", nome)
    nome = _ORAL_RE.sub(" ORAL ", nome)
    nome = _COMPRIMIDO_RE.sub(" ", nome)
    nome = _APRESENTACAO_RE.sub(" ", nome)

    nome = _PONTUACAO_RE.sub(" ", nome)
    nome = _ESPACOS_RE.sub(" ", nome).strip()
    nome = _REPETICAO_RE.sub(r"\1", nome)

    for marcador, original in _PROTEGIDOS_INVERSO.items():
        nome = nome.replace(marcador, original)
    for marcador, original in _RESTAURAR_MARCADORES.items():
        nome = nome.replace(marcador, original)

    nome = _ESPACOS_RE.sub(" ", nome).strip()
    return _EQUIVALENCIAS.get(nome, nome)


# Alias mantido para compatibilidade com o código existente.
agrupar_medicamento = normalizar_medicamento
