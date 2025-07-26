import pandas as pd
from datetime import datetime, timedelta
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import re
import numpy as np

def normalizar_nome(nome):
    if pd.isna(nome):
        return ''
    nome = nome.lower()

    # Equivalentes de forma farmacêutica
    nome = re.sub(r'\b(cápsula|cápsula dura|comprimido|comprimidos|drágea|tablete|cp|cap|tabs?)\b', '', nome)
    nome = re.sub(r'\b(solução|solução oral|sol oral|sol inj|injetável|injeção|ampola)\b', '', nome)
    nome = re.sub(r'\b(suspensão|susp oral|sup)\b', '', nome)
    nome = re.sub(r'\b(água destilada|água para injeção)\b', 'água', nome)

    nome = re.sub(r'\s+', ' ', nome)
    return nome.strip()

def executar_analise(caminho_estoque, caminho_dispensacao, caminho_distribuicao, caminho_saida):
    # Leitura dos dados
    df_estoque = pd.read_excel(caminho_estoque)
    df_dispensacao = pd.read_excel(caminho_dispensacao)
    df_distribuicao = pd.read_excel(caminho_distribuicao)

    # Normalização
    df_estoque['Medicamento/Produto Normalizado'] = df_estoque['Medicamento/Produto'].apply(normalizar_nome)
    df_dispensacao['Medicamento/Produto Normalizado'] = df_dispensacao['Medicamento/Produto'].apply(normalizar_nome)
    df_distribuicao['Medicamento/Produto Normalizado'] = df_distribuicao['Medicamento/Produto'].apply(normalizar_nome)

    # Limpa linhas vazias
    df_estoque.dropna(subset=['Medicamento/Produto'], inplace=True)
    df_dispensacao.dropna(subset=['Medicamento/Produto'], inplace=True)
    df_distribuicao.dropna(subset=['Medicamento/Produto'], inplace=True)

    # Agrupamento de estoque por medicamento (soma os lotes)
    estoque_agg = df_estoque.groupby('Medicamento/Produto Normalizado').agg({
        'Quantidade em Estoque': 'sum',
        'Validade': 'max'
    }).reset_index()

    # Processamento de consumo
    df_dispensacao['Data Dispensação'] = pd.to_datetime(df_dispensacao['Data Dispensação'])
    df_distribuicao['Data Distribuição'] = pd.to_datetime(df_distribuicao['Data Distribuição'])

    df_dispensacao_grouped = df_dispensacao.groupby(['Medicamento/Produto Normalizado', 'Data Dispensação'])['Quantidade Dispensada'].sum().reset_index()
    df_distribuicao_grouped = df_distribuicao.groupby(['Medicamento/Produto Normalizado', 'Data Distribuição'])['Quantidade distribuída (unidades)'].sum().reset_index()

    df_dispensacao_grouped.rename(columns={'Data Dispensação': 'Data', 'Quantidade Dispensada': 'Dispensado'}, inplace=True)
    df_distribuicao_grouped.rename(columns={'Data Distribuição': 'Data', 'Quantidade distribuída (unidades)': 'Distribuido'}, inplace=True)

    df_consumo = pd.merge(df_dispensacao_grouped, df_distribuicao_grouped, on=['Medicamento/Produto Normalizado', 'Data'], how='outer')
    df_consumo[['Dispensado', 'Distribuido']] = df_consumo[['Dispensado', 'Distribuido']].fillna(0)
    df_consumo['Total Consumido'] = df_consumo['Dispensado'] + df_consumo['Distribuido']

    def calcular_consumo(medicamento):
        df_med = df_consumo[df_consumo['Medicamento/Produto Normalizado'] == medicamento].copy()
        df_med = df_med.sort_values('Data')

        df_med['Dias sem falta'] = (df_med['Total Consumido'] > 0).astype(int)
        dias_ativos = df_med['Dias sem falta'].sum()
        total_consumido = df_med['Total Consumido'].sum()

        consumo_mensal = (total_consumido / dias_ativos) * 30 if dias_ativos > 0 else 0
        return pd.Series({
            'Total Consumido': total_consumido,
            'Consumo Mensal Médio Corrigido': round(consumo_mensal, 2),
            'Dias com Falta': (df_med['Total Consumido'] == 0).sum()
        })

    analise = estoque_agg.copy()
    analise = analise.merge(df_consumo['Medicamento/Produto Normalizado'].drop_duplicates(), on='Medicamento/Produto Normalizado', how='left')
    analise = analise.merge(df_consumo.groupby('Medicamento/Produto Normalizado').apply(calcular_consumo).reset_index(), on='Medicamento/Produto Normalizado', how='left')

    # Cálculo da previsão do fim do estoque
    analise['Previsão Fim do Estoque'] = analise.apply(
        lambda row: (datetime.today() + timedelta(days=(row['Quantidade em Estoque'] / row['Consumo Mensal Médio Corrigido']) * 30)).date()
        if pd.notna(row['Consumo Mensal Médio Corrigido']) and row['Consumo Mensal Médio Corrigido'] > 0 else 'FALTA',
        axis=1
    )

    # Preparar resultado final
    resultado = analise[[
        'Medicamento/Produto Normalizado', 'Quantidade em Estoque', 'Validade',
        'Consumo Mensal Médio Corrigido', 'Previsão Fim do Estoque', 'Dias com Falta'
    ]].rename(columns={
        'Medicamento/Produto Normalizado': 'Medicamento',
        'Validade': 'Validade mais distante'
    })

    # Exportar para Excel com destaque de cores
    resultado.to_excel(caminho_saida, index=False)
    wb = load_workbook(caminho_saida)
    ws = wb.active

    vermelho = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    amarelo = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")

    for row in range(2, ws.max_row + 1):
        previsao = ws[f'E{row}'].value
        if isinstance(previsao, datetime):
            dias = (previsao.date() - datetime.today().date()).days
            if dias < 120:
                cor = vermelho if dias < 60 else amarelo
                ws[f'E{row}'].fill = cor
        elif previsao == 'FALTA':
            ws[f'E{row}'].fill = vermelho

    wb.save(caminho_saida)