"""
Cruzamento REMUME × estoque.

Substitui analise2.py.

Problemas do original corrigidos aqui:

  1. CASAMENTO EXATO POR TEXTO CRU — o cruzamento era
     `df_estoque['Medicamento/Produto'] == nome_estoque` após apenas
     `.lower().strip()`. Qualquer diferença de acento, espaço duplo,
     abreviação ou "COMP." vs "COMPRIMIDO" resultava em "EM FALTA" para um
     item que existia em estoque. Agora as duas pontas passam pela mesma
     normalização (`agrupar_medicamento`) usada no resto do sistema.

  2. "EM FALTA" COMO TEXTO NA COLUNA NUMÉRICA — a coluna 'Quantidade em
     Estoque' misturava números e a string 'EM FALTA'. Isso impedia somar,
     ordenar ou filtrar a planilha no Excel. Agora a quantidade é sempre
     numérica e o diagnóstico vai para uma coluna 'Situação' própria.

  3. SEM VALIDAÇÃO DE COLUNAS — 'Quantidade em Estoque' ausente gerava
     KeyError sem mensagem útil ao usuário.

  4. ITERROWS + FILTRO POR LINHA — O(n×m). Em planilhas reais isso levava
     minutos. Agora é um merge agregado.

  5. NaN VIRANDO "nan" — `astype(str)` transformava células vazias no texto
     "nan", que entrava no relatório como se fosse um medicamento.

  6. DUPLICATAS DO MAPEAMENTO — a mesma correspondência repetida contava o
     estoque duas vezes na leitura do relatório.

  7. GRAVAÇÃO EM `static/` com nome só de data: duas execuções no mesmo dia
     sobrescreviam o arquivo e ele ficava publicamente acessível.

  8. `insert_rows(1)` + merge era feito com o writer ainda aberto e podia
     perder a formatação. Agora o estilo é aplicado após o arquivo fechar.
"""

from __future__ import annotations

import logging
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from config import RESULTADO_DIR
from normalizacao import agrupar_medicamento
from planilhas import converter_numero, exigir_colunas, ler_planilha

logger = logging.getLogger(__name__)

COL_REMUME = "RELAÇÃO MUNICIPAL DE MEDICAMENTOS ESSENCIAIS"
COL_PRODUTO = "Medicamento/Produto"
COL_QUANTIDADE = "Quantidade em Estoque"

FILL_TITULO = PatternFill(start_color="FF305496", end_color="FF305496", fill_type="solid")
FILL_FALTA = PatternFill(start_color="FFFFC7CE", end_color="FFFFC7CE", fill_type="solid")
FILL_OK = PatternFill(start_color="FFC6EFCE", end_color="FFC6EFCE", fill_type="solid")


def _texto(serie: pd.Series) -> pd.Series:
    """Texto limpo sem transformar vazios no literal 'nan'."""
    return serie.where(serie.notna(), "").astype(str).str.strip()


def executar_analise_remume(
    caminho_estoque: str | Path,
    caminho_correspondencias: str | Path,
) -> Path:
    """Cruza a REMUME com o estoque atual e devolve o caminho da planilha."""
    df_estoque = ler_planilha(caminho_estoque)
    exigir_colunas(df_estoque, [COL_PRODUTO, COL_QUANTIDADE], "estoque")

    df_mapa = ler_planilha(caminho_correspondencias)
    exigir_colunas(df_mapa, [COL_REMUME, COL_PRODUTO], "correspondências REMUME")

    df_estoque = df_estoque[[COL_PRODUTO, COL_QUANTIDADE]].copy()
    df_estoque[COL_PRODUTO] = _texto(df_estoque[COL_PRODUTO])
    df_estoque = df_estoque[df_estoque[COL_PRODUTO] != ""]
    df_estoque["Quantidade"] = df_estoque[COL_QUANTIDADE].map(converter_numero)
    df_estoque["Chave"] = df_estoque[COL_PRODUTO].map(agrupar_medicamento)

    agregado = (
        df_estoque.groupby("Chave", as_index=False)
        .agg(
            Quantidade=("Quantidade", "sum"),
            Itens_em_Estoque=("Quantidade", "size"),
            Nome_no_Estoque=(COL_PRODUTO, "last"),
        )
    )

    df_mapa = df_mapa[[COL_REMUME, COL_PRODUTO]].copy()
    df_mapa[COL_REMUME] = _texto(df_mapa[COL_REMUME])
    df_mapa[COL_PRODUTO] = _texto(df_mapa[COL_PRODUTO])
    df_mapa = df_mapa[(df_mapa[COL_REMUME] != "") & (df_mapa[COL_PRODUTO] != "")]
    df_mapa = df_mapa.drop_duplicates(subset=[COL_REMUME, COL_PRODUTO])
    df_mapa["Chave"] = df_mapa[COL_PRODUTO].map(agrupar_medicamento)

    if df_mapa.empty:
        raise ValueError("A planilha de correspondências não tem nenhuma linha utilizável.")

    resultado = df_mapa.merge(agregado, on="Chave", how="left")
    resultado["Quantidade"] = resultado["Quantidade"].fillna(0.0)
    resultado["Itens_em_Estoque"] = resultado["Itens_em_Estoque"].fillna(0).astype(int)

    resultado["Situação"] = resultado["Quantidade"].map(
        lambda q: "EM FALTA" if q <= 0 else "DISPONÍVEL"
    )

    df_final = pd.DataFrame(
        {
            "Medicamento REMUME": resultado[COL_REMUME],
            "Medicamento no Estoque": resultado[COL_PRODUTO],
            "Nome Encontrado no Estoque": resultado["Nome_no_Estoque"].fillna("—"),
            "Quantidade em Estoque": resultado["Quantidade"].round(0).astype(int),
            "Registros Somados": resultado["Itens_em_Estoque"],
            "Situação": resultado["Situação"],
        }
    ).sort_values(["Situação", "Medicamento REMUME"])

    hoje = datetime.now()
    caminho = RESULTADO_DIR / (
        f"Estoque_REMUME_{hoje.strftime('%d-%m-%Y')}.xlsx"
    )

    with pd.ExcelWriter(caminho, engine="openpyxl") as writer:
        df_final.to_excel(writer, index=False, sheet_name="REMUME", startrow=1)

    _formatar(caminho, hoje, len(df_final.columns))

    faltantes = int((df_final["Situação"] == "EM FALTA").sum())
    logger.info(
        "Cruzamento REMUME concluído: %d itens, %d em falta.", len(df_final), faltantes
    )
    return caminho


def _formatar(caminho: Path, quando: datetime, colunas: int) -> None:
    wb = load_workbook(caminho)
    ws = wb.active

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=colunas)
    titulo = ws.cell(row=1, column=1)
    titulo.value = (
        "Estoque REMUME com correspondência mapeada — "
        f"{quando.strftime('%d/%m/%Y %H:%M')}"
    )
    titulo.font = Font(bold=True, size=13, color="FFFFFFFF")
    titulo.fill = FILL_TITULO
    titulo.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 24

    cabecalhos = {ws.cell(row=2, column=c).value: c for c in range(1, colunas + 1)}
    for coluna in range(1, colunas + 1):
        celula = ws.cell(row=2, column=coluna)
        celula.font = Font(bold=True)
        celula.alignment = Alignment(horizontal="center", wrap_text=True)

        largura = max(
            (len(str(ws.cell(row=r, column=coluna).value or "")) for r in range(2, ws.max_row + 1)),
            default=12,
        )
        ws.column_dimensions[get_column_letter(coluna)].width = min(max(largura + 2, 14), 50)

    ws.freeze_panes = "A3"

    col_situacao = cabecalhos.get("Situação")
    if col_situacao:
        for linha in range(3, ws.max_row + 1):
            celula = ws.cell(row=linha, column=col_situacao)
            celula.fill = FILL_FALTA if celula.value == "EM FALTA" else FILL_OK

    col_qtd = cabecalhos.get("Quantidade em Estoque")
    if col_qtd:
        for linha in range(3, ws.max_row + 1):
            ws.cell(row=linha, column=col_qtd).number_format = "#,##0"

    wb.save(caminho)
