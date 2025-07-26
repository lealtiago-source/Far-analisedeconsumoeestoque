import pandas as pd
from datetime import datetime, timedelta
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import numpy as np
import os

def executar_analise(caminho_estoque, caminho_dispensacao, caminho_distribuicao):
    try:
        df_estoque = pd.read_excel(caminho_estoque)
        df_disp = pd.read_excel(caminho_dispensacao)
        df_dist = pd.read_excel(caminho_distribuicao)

        # Normalização dos nomes
        def normalizar(nome):
            if pd.isna(nome): return ''
            nome = nome.lower()
            nome = nome.replace('comprimido', '').replace('cápsula', '')
            nome = nome.replace('drágea', '').replace('solução', '')
            nome = nome.replace('água destilada', 'água')
            nome = nome.replace('água para injeção', 'água')
            return ' '.join(nome.split())

        for df in [df_estoque, df_disp, df_dist]:
            df['Medicamento/Produto Normalizado'] = df['Medicamento/Produto'].apply(normalizar)

        # Tratamento de estoques duplicados
        df_estoque = df_estoque.groupby('Medicamento/Produto Normalizado', as_index=False).agg({
            'Quantidade em Estoque': 'sum',
            'Validade': 'max',
            'Medicamento/Produto': 'first'
        })

        # Conversão de datas
        df_disp['Data Dispensação'] = pd.to_datetime(df_disp['Data Dispensação'], errors='coerce')
        df_dist['Data Distribuição'] = pd.to_datetime(df_dist['Data Distribuição'], errors='coerce')

        # Consumo mensal corrigido: só considerar dias com estoque
        def consumo_mensal(df, data_col):
            if df.empty: return 0
            df = df.sort_values(by=data_col)
            total = df['Quantidade Dispensada' if data_col == 'Data Dispensação' else 'Quantidade Distribuída'].sum()
            dias_unicos = df[data_col].dt.date.nunique()
            meses = max(1, dias_unicos / 30)
            return total / meses

        resumo = []

        for med in df_estoque['Medicamento/Produto Normalizado']:
            estoque = df_estoque.loc[df_estoque['Medicamento/Produto Normalizado'] == med, 'Quantidade em Estoque'].values[0]
            validade = df_estoque.loc[df_estoque['Medicamento/Produto Normalizado'] == med, 'Validade'].values[0]
            nome_original = df_estoque.loc[df_estoque['Medicamento/Produto Normalizado'] == med, 'Medicamento/Produto'].values[0]

            disp_med = df_disp[df_disp['Medicamento/Produto Normalizado'] == med]
            dist_med = df_dist[df_dist['Medicamento/Produto Normalizado'] == med]

            total_dispensado = disp_med['Quantidade Dispensada'].sum()
            total_distribuido = dist_med['Quantidade Distribuída'].sum()
            total_consumido = total_dispensado + total_distribuido

            consumo_corrigido_disp = consumo_mensal(disp_med, 'Data Dispensação')
            consumo_corrigido_dist = consumo_mensal(dist_med, 'Data Distribuição')
            consumo_corrigido = consumo_corrigido_disp + consumo_corrigido_dist

            if consumo_corrigido > 0:
                previsao_dias = int(estoque // consumo_corrigido)
                previsao_fim = datetime.today() + timedelta(days=previsao_dias)
                previsao = previsao_fim.strftime('%d/%m/%Y')
                falta = ''
            else:
                previsao = ''
                data_ultima = max(
                    disp_med['Data Dispensação'].max() if not disp_med.empty else datetime.min,
                    dist_med['Data Distribuição'].max() if not dist_med.empty else datetime.min
                )
                if pd.isna(data_ultima):
                    falta = 'FALTA SEM REGISTRO DE USO'
                else:
                    falta = f'FALTA DESDE {data_ultima.strftime("%d/%m/%Y")}'

            resumo.append({
                'Medicamento/Produto': nome_original,
                'Quantidade em Estoque': estoque,
                'Validade': validade,
                'Total Consumido': total_consumido,
                'Consumo Mensal Médio Corrigido': round(consumo_corrigido, 2),
                'Previsão Fim do Estoque': previsao,
                'Situação de Falta': falta
            })

        df_final = pd.DataFrame(resumo)

        # Salva com formatação de cores
        caminho_saida = 'static/resultado_analise.xlsx'
        df_final.to_excel(caminho_saida, index=False)

        # Aplicar cores
        wb = load_workbook(caminho_saida)
        ws = wb.active
        col_idx = {cell.value: idx+1 for idx, cell in enumerate(ws[1])}

        hoje = datetime.today()

        for row in range(2, ws.max_row + 1):
            celula_data = ws.cell(row=row, column=col_idx['Previsão Fim do Estoque'])
            if celula_data.value and isinstance(celula_data.value, str):
                try:
                    data_prev = datetime.strptime(celula_data.value, "%d/%m/%Y")
                    if (data_prev - hoje).days < 120:
                        for col in range(1, ws.max_column + 1):
                            ws.cell(row=row, column=col).fill = PatternFill(start_color="FF9999", fill_type="solid")
                except:
                    continue

        wb.save(caminho_saida)

        return True

    except Exception as e:
        print("Erro na análise:", str(e))
        return False