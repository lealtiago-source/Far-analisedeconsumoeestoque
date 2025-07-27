
import pandas as pd
import os
import re
from datetime import datetime

def normalizar_medicamento(nome):
    if pd.isna(nome): return ''
    nome = str(nome).lower()

    # Palavras a remover (formas farmacêuticas, sais e termos irrelevantes)
    blacklist = [
        'comprimido', 'cápsula', 'solução', 'xarope', 'injeção', 'dura', 'revestido',
        'monoidratado', 'dihidratado', 'triidratado', 'gotas', 'suspensão',
        'cloreto', 'sódico', 'potássico', 'butilbrometo', 'de', 'para', 'uso', 'oral'
    ]
    for termo in blacklist:
        nome = re.sub(rf'\b{termo}\b', '', nome)

    # Extrai concentração (ex: 10 mg, 5mg/ml, 500mcg etc.)
    padrao_conc = re.findall(r'\d+\s?(?:mg|mcg|g|ml|mg/ml|mgml)', nome)
    concentracao = ' '.join(padrao_conc)

    # Remove concentração do nome
    nome = re.sub(r'\d+\s?(mg|mcg|g|ml|mg/ml|mgml)', '', nome)

    # Limpeza final
    nome = re.sub(r'[^\w\s]', '', nome)
    nome = re.sub(r'\s+', ' ', nome).strip()

    return f"{nome} {concentracao}".strip()

def executar_analise_remume(arquivo_estoque, arquivo_remume):
    try:
        df_estoque = pd.read_excel(arquivo_estoque)
        df_remume = pd.read_excel(arquivo_remume)

        # Corrige nomes de colunas
        df_estoque = df_estoque.rename(columns={
            'Lote ': 'Lote',
            'Quantidade em Estoque ': 'Quantidade em Estoque'
        })

        # Normaliza nomes
        df_remume['Normalizado'] = df_remume.iloc[:, 0].apply(normalizar_medicamento)
        df_estoque['Normalizado'] = df_estoque['Medicamento/Produto'].apply(normalizar_medicamento)

        # Agrupamento de estoque
        df_agrupado = df_estoque.groupby('Normalizado').agg({
            'Quantidade em Estoque': 'sum'
        }).reset_index()

        df_agrupado['Presente no REMUME'] = df_agrupado['Normalizado'].isin(df_remume['Normalizado'])

        nomes_originais = df_estoque[['Normalizado', 'Medicamento/Produto']].drop_duplicates('Normalizado', keep='last')
        df_final = df_agrupado.merge(nomes_originais, on='Normalizado', how='left')

        # Ordena com os do REMUME primeiro
        df_final = df_final.sort_values(by='Presente no REMUME', ascending=False)

        # Reorganiza colunas
        df_final = df_final[['Medicamento/Produto', 'Quantidade em Estoque', 'Presente no REMUME']]

        # Garante existência da pasta static
        os.makedirs("static", exist_ok=True)

        # Gera nome do arquivo com data
        hoje = datetime.today().strftime('%d-%m-%Y')
        caminho_saida = os.path.join("static", f"Estoque_REMUME_atualizado_{hoje}.xlsx")

        df_final.to_excel(caminho_saida, index=False)

    except Exception as e:
        print(f"Erro na análise REMUME: {e}")
