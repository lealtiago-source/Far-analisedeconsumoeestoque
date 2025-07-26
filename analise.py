
import pandas as pd
from datetime import datetime, timedelta
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import re
import os

# Cores
FILL_VERMELHO = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')
FILL_ROXO = PatternFill(start_color='B4A7D6', end_color='B4A7D6', fill_type='solid')
FILL_VERDE = PatternFill(start_color='A9D08E', end_color='A9D08E', fill_type='solid')
FILL_AMARELO = PatternFill(start_color='FFD966', end_color='FFD966', fill_type='solid')

def agrupar_equivalentes(nome):
    if pd.isna(nome): return ''
    nome = str(nome).lower()
    nome = re.sub(r'\bcomprimido(s)? revestido(s)?\b', '', nome)
    nome = re.sub(r'\bcomprimido(s)?\b', '', nome)
    nome = re.sub(r'\bcápsula(s)? dura(s)?\b', '', nome)
    nome = re.sub(r'\bcápsula(s)?\b', '', nome)
    nome = re.sub(r'\bdrágea(s)?\b', '', nome)
    nome = re.sub(r'[^\w\s]', '', nome)
    nome = re.sub(r'\s+', ' ', nome)
    return nome.strip()

def executar_analise(arquivo_dispensacao, arquivo_distribuicao, arquivo_estoque):
    try:
        df_disp = pd.read_excel(arquivo_dispensacao,
                                usecols=['Data Dispensação', 'Medicamento/Produto', 'Lote', 'Quantidade Dispensada'])
        df_dist = pd.read_excel(arquivo_distribuicao,
                                usecols=['Data Distribuição', 'Medicamento/Produto', 'Lote', 'Quantidade distribuída (unidades)'])
        df_estoque = pd.read_excel(arquivo_estoque)

        df_disp = df_disp.rename(columns={
            'Data Dispensação': 'Data',
            'Quantidade Dispensada': 'Quantidade'
        })

        df_dist = df_dist.rename(columns={
            'Data Distribuição': 'Data',
            'Quantidade distribuída (unidades)': 'Quantidade'
        })

        df_combinado = pd.concat([df_disp, df_dist], ignore_index=True)
        df_combinado['Fonte'] = ['dispensação'] * len(df_disp) + ['distribuição'] * len(df_dist)

        df_combinado['Medicamento Agrupado'] = df_combinado['Medicamento/Produto'].apply(agrupar_equivalentes)
        df_estoque['Medicamento Agrupado'] = df_estoque['Medicamento/Produto'].apply(agrupar_equivalentes)

        df_combinado['Data'] = pd.to_datetime(df_combinado['Data'], errors='coerce')
        df_estoque['Validade'] = pd.to_datetime(df_estoque['Validade'], errors='coerce')

        df_combinado = df_combinado.sort_values(['Medicamento Agrupado', 'Data', 'Lote'])

        data_minima = df_combinado['Data'].min()
        data_maxima = df_combinado['Data'].max()
        dias_totais = (data_maxima - data_minima).days + 1 if pd.notna(data_minima) else 0

        resumo_consumo = []
        for medicamento, dados in df_combinado.groupby('Medicamento Agrupado'):
            dados = dados.sort_values('Data').reset_index(drop=True)
            dias_falta = 0
            for i in range(1, len(dados)):
                if dados.loc[i - 1, 'Lote'] != dados.loc[i, 'Lote']:
                    diff = (dados.loc[i, 'Data'] - dados.loc[i - 1, 'Data']).days
                    if diff > 15:
                        dias_falta += diff
            consumo_total = dados['Quantidade'].sum()
            dias_validos = max(dias_totais - dias_falta, 1)
            consumo_mensal = round((consumo_total / dias_validos) * 30, 2)
            ultima_data = dados['Data'].max()
            nome_original = dados['Medicamento/Produto'].iloc[-1]
            resumo_consumo.append({
                'Medicamento Agrupado': medicamento,
                'Medicamento': nome_original,
                'Consumo Total': consumo_total,
                'Dias com Falta': dias_falta,
                'Período Analisado (dias)': dias_validos,
                'Consumo Mensal Médio Corrigido': consumo_mensal,
                'Última Dispensação': ultima_data
            })

        resumo = pd.DataFrame(resumo_consumo)

        meds_estoque_unicos = df_estoque[~df_estoque['Medicamento Agrupado'].isin(resumo['Medicamento Agrupado'])]
        if not meds_estoque_unicos.empty:
            for _, row in meds_estoque_unicos.iterrows():
                resumo = pd.concat([resumo, pd.DataFrame([{
                    'Medicamento Agrupado': row['Medicamento Agrupado'],
                    'Medicamento': row['Medicamento/Produto'],
                    'Consumo Total': 0,
                    'Dias com Falta': 0,
                    'Período Analisado (dias)': 0,
                    'Consumo Mensal Médio Corrigido': 0,
                    'Última Dispensação': pd.NaT
                }])], ignore_index=True)

        estoque_val = df_estoque.groupby('Medicamento Agrupado').agg({
            'Quantidade em Estoque': 'sum',
            'Validade': 'min'
        }).reset_index()

        resumo = resumo.merge(estoque_val, on='Medicamento Agrupado', how='left')

        hoje = datetime.today()

        def prever_estoque_ou_falta(row):
            consumo = row['Consumo Mensal Médio Corrigido']
            estoque = row['Quantidade em Estoque']
            if isinstance(estoque, (int, float)) and estoque > 0 and consumo > 0:
                dias_restantes = (estoque / consumo) * 30
                return hoje + timedelta(days=dias_restantes)
            return pd.NaT

        def ajustar_quantidade_estoque(row):
            estoque = row['Quantidade em Estoque']
            if pd.isna(estoque) or estoque == 0:
                if pd.notna(row['Última Dispensação']):
                    return f"FALTA DESDE {row['Última Dispensação'].date().strftime('%d/%m/%Y')}"
                else:
                    return "FALTA"
            return estoque

        resumo['Previsão Fim do Estoque'] = resumo.apply(prever_estoque_ou_falta, axis=1)
        resumo['Quantidade em Estoque'] = resumo.apply(ajustar_quantidade_estoque, axis=1)

        resumo_final = resumo.drop(columns=['Medicamento Agrupado', 'Última Dispensação'])

        caminho_saida = os.path.join("static", "resultado_analise.xlsx")
        with pd.ExcelWriter(caminho_saida, engine='openpyxl') as writer:
            resumo_final.to_excel(writer, index=False)

        # Aplicar formatação
        wb = load_workbook(caminho_saida)
        ws = wb.active

        col_previsao = col_validade = None
        for col in range(1, ws.max_column + 1):
            valor_cab = ws.cell(row=1, column=col).value
            if valor_cab == 'Previsão Fim do Estoque':
                col_previsao = col
            elif valor_cab == 'Validade':
                col_validade = col

        if col_previsao and col_validade:
            limite_alerta = hoje + timedelta(days=120)
            for row in range(2, ws.max_row + 1):
                cell_prev = ws.cell(row=row, column=col_previsao)
                cell_val = ws.cell(row=row, column=col_validade)
                try:
                    previsao = pd.to_datetime(cell_prev.value)
                    validade = pd.to_datetime(cell_val.value)
                    cell_prev.number_format = 'DD/MM/YYYY'
                    cell_val.number_format = 'DD/MM/YYYY'

                    if validade < previsao:
                        cell_val.fill = FILL_ROXO
                    elif validade < previsao - timedelta(days=120):
                        cell_val.fill = FILL_VERDE
                    elif validade >= previsao - timedelta(days=120) and validade >= hoje:
                        cell_val.fill = FILL_AMARELO

                    if previsao < limite_alerta:
                        cell_prev.fill = FILL_VERMELHO
                except:
                    continue

        wb.save(caminho_saida)

    except Exception as e:
        print(f"Erro na análise: {e}")
