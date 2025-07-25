import pandas as pd
from datetime import datetime, timedelta
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import re
import os


def normalizar_nome(nome):
    if pd.isna(nome):
        return ""
    nome = nome.lower()

    # Agrupar equivalentes
    equivalentes = [
        (r'\bcomprimido\b|\bcápsula\b|\bdrágea\b', ''),
        (r'\bsolução oral\b|\bsuspensão oral\b', 'solução'),
        (r'\bágua destilada\b|\bágua para injeção\b', 'água'),
    ]

    for padrao, substituto in equivalentes:
        nome = re.sub(padrao, substituto, nome)

    # Limpar espaços e pontuações duplicadas
    nome = re.sub(r'[^\w\s/%]', '', nome)
    nome = re.sub(r'\s+', ' ', nome).strip()
    return nome


def executar_analise(caminho_dispensacao, caminho_distribuicao, caminho_estoque, caminho_saida):
    try:
        # Carregar os dados
        df_disp = pd.read_excel(caminho_dispensacao)
        df_dist = pd.read_excel(caminho_distribuicao)
        df_estoque = pd.read_excel(caminho_estoque)

        # Remover linhas vazias
        df_disp.dropna(subset=['Medicamento/Produto'], inplace=True)
        df_dist.dropna(subset=['Medicamento/Produto'], inplace=True)
        df_estoque.dropna(subset=['Medicamento/Produto'], inplace=True)

        # Normalizar nomes
        df_disp['Medicamento Normalizado'] = df_disp['Medicamento/Produto'].apply(normalizar_nome)
        df_dist['Medicamento Normalizado'] = df_dist['Medicamento/Produto'].apply(normalizar_nome)
        df_estoque['Medicamento Normalizado'] = df_estoque['Medicamento/Produto'].apply(normalizar_nome)

        # Somar quantidades dispensadas e distribuídas
        disp = df_disp.groupby('Medicamento Normalizado')['Quantidade Dispensada'].sum()
        dist = df_dist.groupby('Medicamento Normalizado')['Quantidade distribuída (unidades)'].sum()

        # Calcular total consumido (dispensado + distribuído)
        total_consumido = disp.add(dist, fill_value=0)

        # Dias cobertos (sem falta)
        dias_disp = df_disp.copy()
        dias_disp['Data Dispensação'] = pd.to_datetime(dias_disp['Data Dispensação'], errors='coerce')
        dias_validos = dias_disp.dropna(subset=['Data Dispensação'])
        dias_por_medicamento = dias_validos.groupby('Medicamento Normalizado')['Data Dispensação'].nunique()

        consumo_medio_corrigido = total_consumido / dias_por_medicamento
        consumo_medio_corrigido = consumo_medio_corrigido.fillna(0).round(2)

        # Estoque atual agrupado por medicamento
        estoque = df_estoque.copy()
        estoque['Validade'] = pd.to_datetime(estoque['Validade'], errors='coerce')
        estoque_agrupado = estoque.groupby('Medicamento Normalizado').agg({
            'Quantidade em Estoque': 'sum',
            'Validade': 'min'
        }).reset_index()

        # Juntar tudo
        resumo = pd.DataFrame({'Total Consumido': total_consumido})
        resumo['Consumo Médio Diário Corrigido'] = consumo_medio_corrigido
        resumo = resumo.reset_index().rename(columns={'index': 'Medicamento Normalizado'})
        resumo = pd.merge(resumo, estoque_agrupado, on='Medicamento Normalizado', how='left')

        # Previsão de fim de estoque
        resumo['Previsão Fim do Estoque'] = resumo.apply(
            lambda row: (datetime.today() + timedelta(days=int(row['Quantidade em Estoque'] / row['Consumo Médio Diário Corrigido'])))
            if row['Consumo Médio Diário Corrigido'] > 0 else 'FALTA',
            axis=1
        )

        # Diagnóstico da falta
        resumo['Análise de Falta'] = resumo['Quantidade em Estoque'].apply(
            lambda q: 'FALTA' if q == 0 or pd.isna(q) else 'OK'
        )

        # Corrigir coluna final de previsão
        resumo['Previsão Fim do Estoque'] = pd.to_datetime(resumo['Previsão Fim do Estoque'], errors='coerce').dt.date

        # Organizar colunas finais
        colunas_finais = [
            'Medicamento Normalizado', 'Total Consumido',
            'Consumo Médio Diário Corrigido', 'Quantidade em Estoque',
            'Validade', 'Previsão Fim do Estoque', 'Análise de Falta'
        ]
        resumo_final = resumo[colunas_finais]

        # Salvar com formatação
        arquivo_saida = os.path.join(caminho_saida, 'analise_consumo_estoque.xlsx')
        resumo_final.to_excel(arquivo_saida, index=False)

        # Colorir com openpyxl
        wb = load_workbook(arquivo_saida)
        ws = wb.active

        for i, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row), start=2):
            qtd = row[3].value  # Quantidade em Estoque
            previsao = row[5].value  # Previsão fim
            analise = row[6].value

            # Cores
            vermelho = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
            amarelo = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')
            verde = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')

            if analise == 'FALTA':
                for cell in row:
                    cell.fill = vermelho
            elif previsao and isinstance(previsao, datetime.date):
                dias_restantes = (previsao - datetime.today().date()).days
                if dias_restantes < 120:
                    for cell in row:
                        cell.fill = amarelo
                else:
                    for cell in row:
                        cell.fill = verde

        wb.save(arquivo_saida)
        print("Análise concluída com sucesso.")

    except Exception as e:
        print(f"Erro na análise: {e}")