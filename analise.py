import pandas as pd
import os
from datetime import datetime, timedelta
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import re

def normalizar_nome(nome):
    if pd.isna(nome):
        return ''
    nome = nome.lower()
    nome = re.sub(r'\b(cápsula|comprimido|dr[áa]gea|solu[çc][aã]o|frasco|ampola|mg/ml|mg|ml|unidade|un)\b', '', nome)
    nome = re.sub(r'\s+', ' ', nome).strip()
    if 'água' in nome:
        return 'água'
    return nome

def agrupar_equivalentes(df, coluna):
    df[coluna] = df[coluna].astype(str).apply(normalizar_nome)
    return df

def executar_analise(arquivo_dispensacao, arquivo_distribuicao, arquivo_estoque):
    try:
        df_disp = pd.read_excel(arquivo_dispensacao)
        df_dist = pd.read_excel(arquivo_distribuicao)
        df_est = pd.read_excel(arquivo_estoque)
        
        col_disp = ['Medicamento/Produto', 'Data Dispensação', 'Quantidade Dispensada']
        col_dist = ['Medicamento/Produto', 'Data Distribuição', 'Quantidade Distribuída']
        col_est = ['Medicamento/Produto', 'Quantidade em Estoque', 'Validade']

        for df, cols, nome in [(df_disp, col_disp, "Dispensação"), (df_dist, col_dist, "Distribuição"), (df_est, col_est, "Estoque")]:
            for col in cols:
                if col not in df.columns:
                    raise Exception(f"Colunas esperadas não encontradas em {nome}: {col}")

        df_disp = df_disp[col_disp].dropna(subset=['Medicamento/Produto'])
        df_dist = df_dist[col_dist].dropna(subset=['Medicamento/Produto'])
        df_est = df_est[col_est].dropna(subset=['Medicamento/Produto'])

        df_disp = agrupar_equivalentes(df_disp, 'Medicamento/Produto')
        df_dist = agrupar_equivalentes(df_dist, 'Medicamento/Produto')
        df_est = agrupar_equivalentes(df_est, 'Medicamento/Produto')

        df_disp['Data Dispensação'] = pd.to_datetime(df_disp['Data Dispensação'], errors='coerce')
        df_dist['Data Distribuição'] = pd.to_datetime(df_dist['Data Distribuição'], errors='coerce')
        df_uso = pd.concat([
            df_disp.rename(columns={'Data Dispensação': 'Data', 'Quantidade Dispensada': 'Quantidade'}),
            df_dist.rename(columns={'Data Distribuição': 'Data', 'Quantidade Distribuída': 'Quantidade'})
        ])
        df_uso = df_uso.dropna(subset=['Data'])

        hoje = datetime.now().date()
        resultado = []

        for medicamento in df_uso['Medicamento/Produto'].unique():
            df_med = df_uso[df_uso['Medicamento/Produto'] == medicamento].copy()
            df_med = df_med.sort_values('Data')

            total = df_med['Quantidade'].sum()
            dias_unicos = df_med['Data'].dt.date.nunique()
            consumo_medio = total / dias_unicos if dias_unicos else 0

            estoque_info = df_est[df_est['Medicamento/Produto'] == medicamento]
            quantidade_estoque = estoque_info['Quantidade em Estoque'].sum() if not estoque_info.empty else 0
            validade = estoque_info['Validade'].max() if not estoque_info.empty else None

            if quantidade_estoque == 0:
                datas = df_med[df_med['Quantidade'] == 0]['Data']
                falta_desde = datas.min().date() if not datas.empty else "Sem dados"
                previsao_fim = f"FALTA DESDE {falta_desde}"
            elif consumo_medio > 0:
                dias_restantes = int(quantidade_estoque // consumo_medio)
                data_fim = hoje + timedelta(days=dias_restantes)
                previsao_fim = data_fim.strftime('%d/%m/%Y')
            else:
                previsao_fim = "Consumo médio zero"

            resultado.append({
                'Medicamento/Produto': medicamento,
                'Quantidade em Estoque': quantidade_estoque if quantidade_estoque > 0 else f"FALTA DESDE {falta_desde}",
                'Validade': validade.strftime('%d/%m/%Y') if validade else '',
                'Consumo Médio Diário Corrigido': round(consumo_medio, 2),
                'Previsão Fim do Estoque': previsao_fim
            })

        resumo_final = pd.DataFrame(resultado)

        caminho_saida = os.path.join("static", "resultado_analise.xlsx")
        resumo_final.to_excel(caminho_saida, index=False)

        wb = load_workbook(caminho_saida)
        ws = wb.active
        red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        for row in range(2, ws.max_row + 1):
            val = ws[f'E{row}'].value
            try:
                if isinstance(val, str) and re.match(r'\d{2}/\d{2}/\d{4}', val):
                    data_fim = datetime.strptime(val, "%d/%m/%Y").date()
                    if (data_fim - hoje).days < 120:
                        for col in "ABCDE":
                            ws[f"{col}{row}"].fill = red_fill
            except:
                continue

        wb.save(caminho_saida)

    except Exception as e:
        print(f"Erro na análise: {e}")
