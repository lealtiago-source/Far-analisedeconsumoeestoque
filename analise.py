import pandas as pd
from datetime import datetime, timedelta
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import re
import os

def normalizar_nome(medicamento):
    if pd.isna(medicamento):
        return ''
    medicamento = medicamento.lower()
    medicamento = re.sub(r'\b(comprimido|cápsula|drágea)\b', '', medicamento)
    medicamento = re.sub(r'\b(água destilada|água para injeção)\b', 'água', medicamento)
    medicamento = re.sub(r'\s+', ' ', medicamento).strip()
    return medicamento

def executar_analise(caminho_dispensacao, caminho_distribuicao, caminho_estoque):
    try:
        df_disp = pd.read_excel(caminho_dispensacao)
        df_dist = pd.read_excel(caminho_distribuicao)
        df_estoque = pd.read_excel(caminho_estoque)

        # Renomear colunas para garantir consistência
        df_disp = df_disp.rename(columns={
            'Data Dispensação': 'Data',
            'Medicamento/Produto': 'Medicamento',
            'Lote': 'Lote',
            'Quantidade Dispensada': 'Quantidade'
        })

        df_dist = df_dist.rename(columns={
            'Data Distribuição': 'Data',
            'Medicamento/Produto': 'Medicamento',
            'Lote': 'Lote',
            'Quantidade distribuída (unidades)': 'Quantidade'
        })

        df_estoque = df_estoque.rename(columns={
            'Medicamento/Produto': 'Medicamento',
            'Validade': 'Validade',
            'Lote': 'Lote',
            'Quantidade em Estoque': 'Quantidade em Estoque'
        })

        # Normalizar
        df_disp['Medicamento Normalizado'] = df_disp['Medicamento'].apply(normalizar_nome)
        df_dist['Medicamento Normalizado'] = df_dist['Medicamento'].apply(normalizar_nome)
        df_estoque['Medicamento Normalizado'] = df_estoque['Medicamento'].apply(normalizar_nome)

        # Combinar dispensação e distribuição
        df_uso = pd.concat([df_disp, df_dist])
        df_uso = df_uso.dropna(subset=['Data', 'Medicamento Normalizado', 'Quantidade'])

        # Agrupar por medicamento e mês
        df_uso['AnoMes'] = df_uso['Data'].dt.to_period('M')
        consumo_mensal = df_uso.groupby(['Medicamento Normalizado', 'AnoMes'])['Quantidade'].sum().reset_index()

        # Dias com consumo no mês
        dias_validos = df_uso.groupby(['Medicamento Normalizado', 'AnoMes'])['Data'].nunique().reset_index()
        dias_validos = dias_validos.rename(columns={'Data': 'Dias'})

        consumo_corrigido = consumo_mensal.merge(dias_validos, on=['Medicamento Normalizado', 'AnoMes'])
        consumo_corrigido['Consumo Corrigido'] = consumo_corrigido['Quantidade'] / consumo_corrigido['Dias']

        # Média mensal corrigida (últimos 3 meses, se possível)
        media_corrigida = consumo_corrigido.groupby('Medicamento Normalizado')['Consumo Corrigido'].mean().reset_index()
        media_corrigida = media_corrigida.rename(columns={'Consumo Corrigido': 'Consumo Mensal Médio Corrigido'})

        # Estoque atual por medicamento
        df_estoque['Quantidade em Estoque'] = df_estoque['Quantidade em Estoque'].fillna(0)
        estoque_atual = df_estoque.groupby('Medicamento Normalizado', as_index=False).agg({
            'Quantidade em Estoque': 'sum',
            'Validade': 'max'
        })

        # Mesclar com consumo
        resultado = pd.merge(estoque_atual, media_corrigida, on='Medicamento Normalizado', how='left')

        # Previsão de fim do estoque
        def calcular_previsao(row):
            try:
                if row['Consumo Mensal Médio Corrigido'] > 0:
                    dias = int(row['Quantidade em Estoque'] / (row['Consumo Mensal Médio Corrigido'] / 30))
                    return (datetime.now() + timedelta(days=dias)).date()
                else:
                    return 'Consumo nulo'
            except:
                return 'Erro cálculo'

        resultado['Previsão Fim do Estoque'] = resultado.apply(calcular_previsao, axis=1)

        # Análise de falta
        def analisar_falta(med):
            datas = df_uso[df_uso['Medicamento Normalizado'] == med]['Data']
            dias = pd.date_range(start=datas.min(), end=datas.max())
            usados = datas.dt.date.unique()
            faltantes = [dia for dia in dias.date if dia not in usados]
            if faltantes:
                return f"Faltou em {len(faltantes)} dias"
            return "Sem falta"

        resultado['Análise de Falta'] = resultado['Medicamento Normalizado'].apply(analisar_falta)

        # Reorganizar colunas
        resultado_final = resultado[[
            'Medicamento Normalizado',
            'Quantidade em Estoque',
            'Consumo Mensal Médio Corrigido',
            'Previsão Fim do Estoque',
            'Validade',
            'Análise de Falta'
        ]].copy()

        resultado_final = resultado_final.rename(columns={
            'Medicamento Normalizado': 'Medicamento/Produto'
        })

        # Salvar Excel com destaque
        arquivo_saida = 'static/resultado_analise.xlsx'
        resultado_final.to_excel(arquivo_saida, index=False)

        wb = load_workbook(arquivo_saida)
        ws = wb.active

        red_fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')

        for row in range(2, ws.max_row + 1):
            val = ws[f'C{row}'].value  # Consumo Médio
            estoque = ws[f'B{row}'].value
            if isinstance(val, (int, float)) and val > 0:
                dias = estoque / (val / 30)
                if dias < 120:
                    for col in range(1, ws.max_column + 1):
                        ws.cell(row=row, column=col).fill = red_fill

        wb.save(arquivo_saida)
        return True, None

    except Exception as e:
        return False, f'Erro na análise: {e}'