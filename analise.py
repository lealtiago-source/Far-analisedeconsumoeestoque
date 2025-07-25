import pandas as pd
from datetime import datetime, timedelta

def executar_analise(caminho_dispensacao, caminho_distribuicao, caminho_estoque, caminho_saida):
    try:
        # Lê a primeira (e única) planilha de cada arquivo
        df_disp = pd.read_excel(caminho_dispensacao, sheet_name=0)
        df_dist = pd.read_excel(caminho_distribuicao, sheet_name=0)
        df_estoque = pd.read_excel(caminho_estoque, sheet_name=0)

        # Normalização básica
        df_disp = df_disp.dropna(subset=['Medicamento/Produto'])
        df_dist = df_dist.dropna(subset=['Medicamento/Produto'])
        df_estoque = df_estoque.dropna(subset=['Medicamento/Produto'])

        # Normaliza os nomes dos medicamentos
        def normalizar(nome):
            nome = str(nome).lower()
            nome = nome.replace("comprimido", "").replace("cápsula", "")
            nome = nome.replace("drágea", "").replace("mg", "").strip()
            return nome

        df_disp['Produto Normalizado'] = df_disp['Medicamento/Produto'].map(normalizar)
        df_dist['Produto Normalizado'] = df_dist['Medicamento/Produto'].map(normalizar)
        df_estoque['Produto Normalizado'] = df_estoque['Medicamento/Produto'].map(normalizar)

        # Agrupa as quantidades
        dispensado = df_disp.groupby('Produto Normalizado')['Quantidade Dispensada'].sum()
        distribuido = df_dist.groupby('Produto Normalizado')['Quantidade Distribuída'].sum()
        estoque = df_estoque.groupby('Produto Normalizado')['Quantidade em Estoque'].sum()

        # Junta tudo
        df_resultado = pd.concat([dispensado, distribuido, estoque], axis=1)
        df_resultado.columns = ['Total Dispensado', 'Total Distribuído', 'Quantidade em Estoque']
        df_resultado = df_resultado.fillna(0)

        # Previsão de fim do estoque
        df_resultado['Dias de Consumo'] = (df_resultado['Total Dispensado'] + df_resultado['Total Distribuído']) / 30
        df_resultado['Previsão Fim do Estoque'] = df_resultado.apply(
            lambda row: (datetime.today() + timedelta(days=(row['Quantidade em Estoque'] / row['Dias de Consumo'] * 30)))
            if row['Dias de Consumo'] > 0 else 'FALTA DESDE ' + datetime.today().strftime('%d/%m/%Y'),
            axis=1
        )

        # Salva o arquivo
        df_resultado.reset_index().to_excel(caminho_saida, index=False)
    
    except Exception as e:
        print(f"Erro na análise: {e}")