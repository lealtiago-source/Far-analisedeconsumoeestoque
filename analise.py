import pandas as pd
from datetime import datetime, timedelta
import re
from openpyxl import load_workbook
from openpyxl.styles import PatternFill


def normalizar_nome(nome):
    nome = str(nome).lower()
    nome = re.sub(r'\b(comprimido|cápsula|drágea|comprimidos|cápsulas)\b', '', nome)
    nome = re.sub(r'\bágua (destilada|para injeção)\b', 'água', nome)
    nome = re.sub(r'\s+', ' ', nome).strip()
    return nome


def executar_analise(caminho_estoque, caminho_dispensacao, caminho_distribuicao, caminho_saida):
    # Leitura dos dados
    df_estoque = pd.read_excel(caminho_estoque)
    df_dispensacao = pd.read_excel(caminho_dispensacao)
    df_distribuicao = pd.read_excel(caminho_distribuicao)

    # Padronização de colunas
    df_estoque = df_estoque[['Medicamento/Produto', 'Validade', 'Lote', 'Quantidade em Estoque']].copy()
    df_dispensacao = df_dispensacao[['Data Dispensação', 'Medicamento/Produto', 'Lote', 'Quantidade Dispensada']].copy()
    df_distribuicao = df_distribuicao[['Data Distribuição', 'Medicamento/Produto', 'Lote', 'Quantidade distribuída (unidades)']].copy()

    # Normalização de nomes
    df_estoque['Nome Normalizado'] = df_estoque['Medicamento/Produto'].apply(normalizar_nome)
    df_dispensacao['Nome Normalizado'] = df_dispensacao['Medicamento/Produto'].apply(normalizar_nome)
    df_distribuicao['Nome Normalizado'] = df_distribuicao['Medicamento/Produto'].apply(normalizar_nome)

    # Remover linhas vazias
    df_estoque = df_estoque[df_estoque['Medicamento/Produto'].notna()]
    df_dispensacao = df_dispensacao[df_dispensacao['Medicamento/Produto'].notna()]
    df_distribuicao = df_distribuicao[df_distribuicao['Medicamento/Produto'].notna()]

    # Conversão de datas
    df_dispensacao['Data Dispensação'] = pd.to_datetime(df_dispensacao['Data Dispensação'], errors='coerce')
    df_distribuicao['Data Distribuição'] = pd.to_datetime(df_distribuicao['Data Distribuição'], errors='coerce')
    hoje = pd.to_datetime(datetime.today().date())

    # Soma da quantidade por medicamento normalizado
    dispensacao_agrupada = df_dispensacao.groupby('Nome Normalizado').agg(
        Quantidade_Dispensada_Total=('Quantidade Dispensada', 'sum'),
        Dias_Dispensados=('Data Dispensação', lambda x: x.nunique())
    ).reset_index()

    distribuicao_agrupada = df_distribuicao.groupby('Nome Normalizado').agg(
        Quantidade_Distribuida_Total=('Quantidade distribuída (unidades)', 'sum')
    ).reset_index()

    estoque_agrupado = df_estoque.groupby('Nome Normalizado').agg(
        Quantidade_Estoque=('Quantidade em Estoque', 'sum'),
        Validade_Mais_Proxima=('Validade', 'min')
    ).reset_index()

    # Combina as tabelas
    df_final = estoque_agrupado.merge(dispensacao_agrupada, on='Nome Normalizado', how='left')
    df_final = df_final.merge(distribuicao_agrupada, on='Nome Normalizado', how='left')

    df_final.fillna(0, inplace=True)

    # Consumo médio diário considerando apenas dias com dispensação
    df_final['Consumo Médio Diário'] = df_final.apply(
        lambda row: row['Quantidade_Dispensada_Total'] / row['Dias_Dispensados']
        if row['Dias_Dispensados'] > 0 else 0, axis=1
    )

    # Previsão fim do estoque
    df_final['Previsão Fim do Estoque'] = df_final.apply(
        lambda row: (hoje + timedelta(days=row['Quantidade_Estoque'] / row['Consumo Médio Diário']))
        if row['Consumo Médio Diário'] > 0 else 'FALTA DESDE ' + hoje.strftime('%d/%m/%Y'),
        axis=1
    )

    # Reorganizar colunas para saída
    df_final = df_final[[
        'Nome Normalizado',
        'Quantidade_Estoque',
        'Quantidade_Dispensada_Total',
        'Quantidade_Distribuida_Total',
        'Consumo Médio Diário',
        'Previsão Fim do Estoque',
        'Validade_Mais_Proxima'
    ]]

    # Salvar em Excel
    df_final.to_excel(caminho_saida, index=False)

    # Formatação (destaque em vermelho se previsão < 4 meses)
    wb = load_workbook(caminho_saida)
    ws = wb.active

    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

    for row in range(2, ws.max_row + 1):
        valor = ws[f"F{row}"].value
        if isinstance(valor, datetime):
            if valor <= hoje + timedelta(days=120):  # menos de 4 meses
                ws[f"F{row}"].fill = red_fill
        elif isinstance(valor, str) and valor.startswith("FALTA DESDE"):
            ws[f"F{row}"].fill = red_fill

    wb.save(caminho_saida)