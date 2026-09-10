"""
Cálculo de consumo mensal médio corrigido — módulo único.

Antes o mesmo cálculo estava copiado em analise.py, analise_orcamento.py e
(parcialmente) analise_compra.py, cada cópia com uma variação do bug:

  - `dias_totais` global do dataset em vez do período do próprio medicamento
    (subestimava itens que entraram na rede no último mês);
  - `dias_validos = max(dias - falta, 1)`, que ao chegar em 1 dia inflava o
    consumo mensal em até 30x;
  - `dias_falta += diff`, contando o intervalo inteiro como falta em vez do
    excedente sobre o intervalo normal de reposição.

Agora existe uma implementação só, usada por todas as análises.
"""

from __future__ import annotations

import pandas as pd

from config import (
    DIAS_MINIMOS_DESABASTECIMENTO,
    DIAS_MINIMOS_PERIODO,
    PROPORCAO_MINIMA_PERIODO_VALIDO,
)


def dias_de_falta(dados: pd.DataFrame) -> int:
    """
    Dias sem abastecimento inferidos pela troca de lote.

    Conta apenas o EXCEDENTE sobre o intervalo considerado normal, e não o
    intervalo inteiro.
    """
    total = 0
    anterior = None
    for _, linha in dados.iterrows():
        if anterior is not None and anterior["Lote"] != linha["Lote"]:
            if pd.notna(linha["Data"]) and pd.notna(anterior["Data"]):
                intervalo = (linha["Data"] - anterior["Data"]).days
                if intervalo > DIAS_MINIMOS_DESABASTECIMENTO:
                    total += intervalo - DIAS_MINIMOS_DESABASTECIMENTO
        anterior = linha
    return total


def calcular_consumo(
    df: pd.DataFrame,
    coluna_valor: str | None = None,
) -> pd.DataFrame:
    """
    Consumo mensal médio por 'Medicamento Agrupado'.

    Espera as colunas: Medicamento Agrupado, Medicamento/Produto, Lote,
    Data, Quantidade. Se `coluna_valor` for informada, agrega também o valor
    unitário médio ponderado pela quantidade.
    """
    registros = []

    for medicamento, dados in df.groupby("Medicamento Agrupado", sort=False):
        dados = dados.sort_values("Data")
        datas = dados["Data"].dropna()
        if datas.empty:
            continue

        dias_periodo = (datas.max() - datas.min()).days + 1
        amostra_insuficiente = dias_periodo < DIAS_MINIMOS_PERIODO
        dias_periodo = max(dias_periodo, DIAS_MINIMOS_PERIODO)

        falta = dias_de_falta(dados)
        piso = max(int(dias_periodo * PROPORCAO_MINIMA_PERIODO_VALIDO), 1)
        dias_validos = max(dias_periodo - falta, piso)

        consumo_total = float(dados["Quantidade"].sum())
        consumo_mensal = round((consumo_total / dias_validos) * 30, 2)

        registro = {
            "Medicamento Agrupado": medicamento,
            "Medicamento": dados["Medicamento/Produto"].iloc[-1],
            "Consumo Total": consumo_total,
            "Dias com Falta": falta,
            "Período Analisado (dias)": dias_validos,
            "Consumo Mensal Médio Corrigido": consumo_mensal,
            "Amostra Insuficiente": "SIM" if amostra_insuficiente else "",
            "Última Dispensação": datas.max(),
        }

        if coluna_valor and coluna_valor in dados.columns:
            valores = dados[coluna_valor].astype(float)
            pesos = dados["Quantidade"].astype(float)
            # Média PONDERADA pela quantidade. A média simples anterior dava
            # o mesmo peso a um lote de 10 unidades e a um de 10.000.
            if pesos.sum() > 0 and valores.notna().any():
                mascara = valores.notna()
                registro["Valor Unitário Médio (R$)"] = round(
                    float((valores[mascara] * pesos[mascara]).sum() / pesos[mascara].sum()), 4
                )
            else:
                registro["Valor Unitário Médio (R$)"] = round(float(valores.mean() or 0), 4)

        registros.append(registro)

    return pd.DataFrame(registros)
