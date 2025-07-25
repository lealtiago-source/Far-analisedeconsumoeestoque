import pandas as pd
from datetime import datetime, timedelta
import os

def executar_analise(caminho_dispensacao, caminho_distribuicao, caminho_estoque, caminho_saida):
    try:
        # Lê a primeira (única) planilha de cada arquivo
        df_disp = pd.read_excel(caminho_dispensacao, sheet_name=0)
        df_dist = pd.read_excel(caminho_distribuicao, sheet_name=0)
        df_estoque = pd.read_excel(caminho_estoque, sheet_name=0)

        # Verifica colunas (ajuste conforme seu arquivo)
        # Se as colunas forem diferentes, ajuste aqui!
        # Exemplo: 'Quantidade Dispensada' pode ser 'Quantidade Dispensada' ou 'Quantidade Dispensada'
        col_disp = 'Quantidade Dispensada' if 'Quantidade Dispensada' in df_disp.columns else None
        col_dist = 'Quantidade Distribuída' if 'Quantidade Distribuída' in df_dist.columns else None
        col_estoque = 'Quantidade em Estoque' if 'Quantidade em Estoque' in df_estoque.columns else None

        if not all([col_disp, col_dist, col_estoque]):
            raise Exception("Colunas esperadas não encontradas nos arquivos.")

        # Remove linhas sem medicamento
        df_disp = df_disp.dropna(subset=['Medicamento/Produto'])
        df_dist = df_dist.dropna(subset=['Medicamento/Produto'])
        df_estoque = df_estoque.dropna(subset=['Medicamento/Produto'])

        # Função para normalizar nomes (ajuste se quiser)
        def normalizar(nome):
            nome = str(nome).lower()
            for palavra in ["comprimido", "cápsula", "drágea", "mg"]:
                nome = nome.replace(palavra, "")
            return nome.strip()

        df_disp['Produto Normalizado'] = df_disp['Medicamento/Produto'].map(normalizar)
        df_dist['Produto Normalizado'] = df_dist['Medicamento/Produto'].map(normalizar)
        df_estoque['Produto Normalizado'] = df_estoque['Medicamento/Produto'].map(normalizar)

        # Agrupa quantidades
        dispensado = df_disp.groupby('Produto Normalizado')[col_disp].sum()
        distribuido = df_dist.groupby('Produto Normalizado')[col_dist].sum()
        estoque = df_estoque.groupby('Produto Normalizado')[col_estoque].sum()

        # Junta tudo
        df_resultado = pd.concat([dispensado, distribuido, estoque], axis=1)
        df_resultado.columns = ['Total Dispensado', 'Total Distribuído', 'Quantidade em Estoque']
        df_resultado = df_resultado.fillna(0)

        # Previsão fim do estoque
        df_resultado['Dias de Consumo'] = (df_resultado['Total Dispensado'] + df_resultado['Total Distribuído']) / 30
        df_resultado['Previsão Fim do Estoque'] = df_resultado.apply(
            lambda row: (datetime.today() + timedelta(days=(row['Quantidade em Estoque'] / row['Dias de Consumo'] * 30)))
            if row['Dias de Consumo'] > 0 else 'FALTA DESDE ' + datetime.today().strftime('%d/%m/%Y'),
            axis=1
        )

        # Cria pasta se não existir
        os.makedirs(os.path.dirname(caminho_saida), exist_ok=True)

        # Salva resultado
        df_resultado.reset_index().to_excel(caminho_saida, index=False)

    except Exception as e:
        print(f"Erro na análise: {e}")
        raise  # Para o erro subir ao Flask e exibir no browser
