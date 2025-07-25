# script.py

import pandas as pd

def processar_planilhas(arq1, arq2, arq3, caminho_saida):
    # Exemplo básico para gerar um arquivo de saída
    df1 = pd.read_excel(arq1)
    df2 = pd.read_excel(arq2)
    df3 = pd.read_excel(arq3)

    resultado = pd.concat([df1, df2, df3])
    resultado.to_excel(caminho_saida, index=False)
