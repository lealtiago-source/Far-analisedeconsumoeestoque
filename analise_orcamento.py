import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import re
import os

# Cores
FILL_VERDE = PatternFill(start_color='A9D08E', end_color='A9D08E', fill_type='solid')
FILL_AMARELO = PatternFill(start_color='FFD966', end_color='FFD966', fill_type='solid')
FILL_VERMELHO = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')

def agrupar_equivalentes(nome):
    """Normaliza nomes de medicamentos, agrupando formulações equivalentes."""
    if pd.isna(nome): 
        return ''
    nome = str(nome).lower()
    nome = re.sub(r'\bcomprimido(s)? revestido(s)?\b', '', nome)
    nome = re.sub(r'\bcomprimido(s)?\b', '', nome)
    nome = re.sub(r'\bcápsula(s)? dura(s)?\b', '', nome)
    nome = re.sub(r'\bcápsula(s)?\b', '', nome)
    nome = re.sub(r'\bdrágea(s)?\b', '', nome)
    nome = re.sub(r'[^\w\s]', '', nome)
    nome = re.sub(r'\s+', ' ', nome)
    return nome.strip()

def executar_analise_orcamento(arquivo_dispensacao, arquivo_distribuicao):
    try:
        # --- Leitura das planilhas ---
        df_disp = pd.read_excel(arquivo_dispensacao)
        df_disp.columns = df_disp.columns.str.strip()
        df_disp = df_disp[['Data Dispensação', 'Medicamento/Produto', 'Lote', 
                           'Quantidade Dispensada', 'Valor Unitário']]

        df_dist = pd.read_excel(arquivo_distribuicao)
        df_dist.columns = df_dist.columns.str.strip()
        df_dist = df_dist[['Data Distribuição', 'Medicamento/Produto', 'Lote', 
                           'Quantidade distribuída (unidades)', 'Valor unitário']]

        # --- Padroniza colunas ---
        df_disp = df_disp.rename(columns={
            'Data Dispensação': 'Data',
            'Quantidade Dispensada': 'Quantidade',
            'Valor Unitário': 'Valor Unitário'
        })
        df_dist = df_dist.rename(columns={
            'Data Distribuição': 'Data',
            'Quantidade distribuída (unidades)': 'Quantidade',
            'Valor unitário': 'Valor Unitário'
        })

        # --- Adiciona origem e une tudo ---
        df_disp['Fonte'] = 'Dispensação'
        df_dist['Fonte'] = 'Distribuição'
        df_combinado = pd.concat([df_disp, df_dist], ignore_index=True)

        # --- Normaliza nomes ---
        df_combinado['Medicamento Agrupado'] = df_combinado['Medicamento/Produto'].apply(agrupar_equivalentes)
        df_combinado['Data'] = pd.to_datetime(df_combinado['Data'], errors='coerce')

        df_combinado = df_combinado.sort_values(['Medicamento Agrupado', 'Data', 'Lote'])

        # --- Calcula o período total ---
        data_minima = df_combinado['Data'].min()
        data_maxima = df_combinado['Data'].max()
        dias_totais = (data_maxima - data_minima).days + 1 if pd.notna(data_minima) else 0

        # --- Análise por medicamento ---
        resumo_consumo = []
        for medicamento, dados in df_combinado.groupby('Medicamento Agrupado'):
            dados = dados.sort_values('Data').reset_index(drop=True)

            # Calcula dias de falta (como no código original)
            dias_falta = 0
            for i in range(1, len(dados)):
                if dados.loc[i - 1, 'Lote'] != dados.loc[i, 'Lote']:
                    diff = (dados.loc[i, 'Data'] - dados.loc[i - 1, 'Data']).days
                    if diff > 15:
                        dias_falta += diff

            consumo_total = dados['Quantidade'].sum()
            dias_validos = max(dias_totais - dias_falta, 1)
            consumo_mensal_corrigido = round((consumo_total / dias_validos) * 30, 2)

            valor_medio_unitario = dados['Valor Unitário'].mean() if 'Valor Unitário' in dados else 0
            consumo_anual_previsto = consumo_mensal_corrigido * 12
            custo_anual_previsto = round(consumo_anual_previsto * valor_medio_unitario, 2)

            resumo_consumo.append({
                'Medicamento Agrupado': medicamento,
                'Medicamento': dados['Medicamento/Produto'].iloc[-1],
                'Consumo Total': consumo_total,
                'Dias com Falta': dias_falta,
                'Período Analisado (dias)': dias_validos,
                'Consumo Mensal Médio Corrigido': consumo_mensal_corrigido,
                'Consumo Anual Previsto': consumo_anual_previsto,
                'Valor Unitário Médio (R$)': round(valor_medio_unitario, 4),
                'Custo Anual Previsto (R$)': custo_anual_previsto
            })

        resumo = pd.DataFrame(resumo_consumo)

        # --- Salva o resultado ---
        caminho_saida = os.path.join("static", "resultado_analise_orcamento.xlsx")
        with pd.ExcelWriter(caminho_saida, engine='openpyxl') as writer:
            resumo.to_excel(writer, index=False, sheet_name='Orçamento')

        # --- Estilo do Excel ---
        wb = load_workbook(caminho_saida)
        ws = wb.active

        # Destaques visuais
        col_custo = None
        for col in range(1, ws.max_column + 1):
            cab = ws.cell(row=1, column=col).value
            if cab == 'Custo Anual Previsto (R$)':
                col_custo = col
                break

        if col_custo:
            for row in range(2, ws.max_row + 1):
                valor = ws.cell(row=row, column=col_custo).value
                if valor is not None:
                    if valor > 100000:
                        ws.cell(row=row, column=col_custo).fill = FILL_VERMELHO
                    elif valor > 50000:
                        ws.cell(row=row, column=col_custo).fill = FILL_AMARELO
                    else:
                        ws.cell(row=row, column=col_custo).fill = FILL_VERDE

        wb.save(caminho_saida)
        print("✅ Análise de orçamento concluída com sucesso!")
        return caminho_saida

    except Exception as e:
        print(f"❌ Erro na análise de orçamento: {e}")
        raise e
