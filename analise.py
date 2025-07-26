import pandas as pd
from datetime import datetime, timedelta
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import re

def normalizar_nome(nome):
    if pd.isna(nome):
        return ''
    nome = nome.lower()

    nome = re.sub(r'\b(comprimido|cápsula|cápsula dura|drágea)\b', '', nome)
    nome = re.sub(r'\b(água destilada|água para injeção)\b', 'água', nome)
    nome = re.sub(r'[^a-z0-9áéíóúãõç%\s\/.,()-]', '', nome)
    nome = re.sub(r'\s+', ' ', nome).strip()

    return nome

def executar_analise(dispensacao_path, distribuicao_path, estoque_path, _saida_ignorado=None):
    try:
        # Leitura
        df_disp = pd.read_excel(dispensacao_path)
        df_dist = pd.read_excel(distribuicao_path)
        df_estoque = pd.read_excel(estoque_path)

        # Colunas obrigatórias
        for df, cols in [(df_disp, ['Medicamento/Produto', 'Data Dispensação', 'Quantidade Dispensada']),
                         (df_dist, ['Medicamento/Produto', 'Data Distribuição', 'Quantidade Distribuída']),
                         (df_estoque, ['Medicamento/Produto', 'Quantidade em Estoque', 'Validade'])]:
            if not all(c in df.columns for c in cols):
                raise Exception(f"Colunas esperadas não encontradas: {cols}")

        # Remover linhas com produto vazio
        df_disp = df_disp[df_disp['Medicamento/Produto'].notna()]
        df_dist = df_dist[df_dist['Medicamento/Produto'].notna()]
        df_estoque = df_estoque[df_estoque['Medicamento/Produto'].notna()]

        # Normalizar nomes
        df_disp['Produto Normalizado'] = df_disp['Medicamento/Produto'].apply(normalizar_nome)
        df_dist['Produto Normalizado'] = df_dist['Medicamento/Produto'].apply(normalizar_nome)
        df_estoque['Produto Normalizado'] = df_estoque['Medicamento/Produto'].apply(normalizar_nome)

        # Agrupar dispensação
        df_disp['Data Dispensação'] = pd.to_datetime(df_disp['Data Dispensação'], errors='coerce')
        consumo = df_disp.groupby(['Produto Normalizado', df_disp['Data Dispensação'].dt.to_period('M')])['Quantidade Dispensada'].sum().reset_index()
        consumo['Data Dispensação'] = consumo['Data Dispensação'].dt.to_timestamp()

        # Agrupar distribuição
        df_dist['Data Distribuição'] = pd.to_datetime(df_dist['Data Distribuição'], errors='coerce')
        dist = df_dist.groupby(['Produto Normalizado', df_dist['Data Distribuição'].dt.to_period('M')])['Quantidade Distribuída'].sum().reset_index()
        dist['Data Distribuição'] = dist['Data Distribuição'].dt.to_timestamp()

        # Calcular consumo total por mês (dispensado + distribuído)
        meses = sorted(set(consumo['Data Dispensação'].dt.to_period('M')).union(set(dist['Data Distribuição'].dt.to_period('M'))))
        consumo_total = []
        for produto in sorted(set(consumo['Produto Normalizado']).union(dist['Produto Normalizado'])):
            serie = []
            for mes in meses:
                val_disp = consumo[(consumo['Produto Normalizado'] == produto) & (consumo['Data Dispensação'].dt.to_period('M') == mes)]['Quantidade Dispensada'].sum()
                val_dist = dist[(dist['Produto Normalizado'] == produto) & (dist['Data Distribuição'].dt.to_period('M') == mes)]['Quantidade Distribuída'].sum()
                total = val_disp + val_dist
                serie.append((produto, mes.to_timestamp(), total))
            consumo_total.extend(serie)

        df_total = pd.DataFrame(consumo_total, columns=['Produto Normalizado', 'Mês', 'Total Consumido'])
        df_total = df_total[df_total['Total Consumido'] > 0]

        # Consumo médio corrigido (média apenas dos meses com consumo)
        media_corrigida = df_total.groupby('Produto Normalizado')['Total Consumido'].agg(['mean', 'count']).reset_index()
        media_corrigida.columns = ['Produto Normalizado', 'Consumo Médio Corrigido', 'Meses com Consumo']

        # Dias sem consumo
        analise_falta = df_total.groupby('Produto Normalizado').apply(
            lambda x: (
                (pd.date_range(x['Mês'].min(), x['Mês'].max(), freq='MS').difference(x['Mês'])).strftime('%m/%Y').tolist()
            )
        ).reset_index(name='Meses Sem Consumo')

        # Estoque agrupado
        df_estoque['Quantidade em Estoque'] = pd.to_numeric(df_estoque['Quantidade em Estoque'], errors='coerce').fillna(0)
        df_estoque['Validade'] = pd.to_datetime(df_estoque['Validade'], errors='coerce')
        estoque_agrupado = df_estoque.groupby('Produto Normalizado').agg({
            'Quantidade em Estoque': 'sum',
            'Validade': 'min',
            'Medicamento/Produto': 'first'
        }).reset_index()

        # Unir tudo
        resumo = estoque_agrupado.merge(media_corrigida, on='Produto Normalizado', how='left')
        resumo = resumo.merge(analise_falta, on='Produto Normalizado', how='left')

        # Previsão de término
        resumo['Previsão Fim do Estoque'] = resumo.apply(
            lambda row: (
                (datetime.today() + timedelta(days=int(row['Quantidade em Estoque'] / row['Consumo Médio Corrigido'] * 30)))
                if row['Consumo Médio Corrigido'] > 0 else pd.NaT
            ), axis=1
        )

        # Filtros e ajustes
        resumo['Consumo Médio Corrigido'] = resumo['Consumo Médio Corrigido'].round(2)
        resumo['Total Consumido'] = df_total.groupby('Produto Normalizado')['Total Consumido'].sum().reset_index(drop=True)

        # Ordenar colunas
        colunas = [
            'Medicamento/Produto', 'Quantidade em Estoque', 'Validade',
            'Consumo Médio Corrigido', 'Previsão Fim do Estoque',
            'Meses Sem Consumo', 'Total Consumido'
        ]
        resultado = resumo[colunas].copy()

        # Salvar no Excel
        caminho_saida = 'static/resultado_analise.xlsx'
        resultado.to_excel(caminho_saida, index=False)

        # Adicionar cores
        wb = load_workbook(caminho_saida)
        ws = wb.active
        hoje = datetime.today()

        for row in range(2, ws.max_row + 1):
            celula = ws[f'E{row}']
            try:
                data = celula.value
                if isinstance(data, datetime):
                    delta = (data - hoje).days
                    if delta < 120:
                        celula.fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")  # vermelho
                    elif 120 <= delta <= 180:
                        celula.fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")  # amarelo
                    else:
                        celula.fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")  # verde
            except:
                continue

        wb.save(caminho_saida)

    except Exception as e:
        print(f"Erro na análise: {e}")
