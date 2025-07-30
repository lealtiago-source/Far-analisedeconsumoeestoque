import pandas as pd
from datetime import datetime
import os
from openpyxl.styles import Font
from openpyxl import load_workbook

def analise2(caminho_estoque, caminho_remume):
    df_estoque = pd.read_excel(caminho_estoque)
    df_remume = pd.read_excel(caminho_remume)

    col_remume = 'RELAÇÃO MUNICIPAL DE MEDICAMENTOS ESSENCIAIS'
    resultado = []

    # Deixar todos em minúsculo para facilitar comparações
    df_estoque['Medicamento/Produto'] = df_estoque['Medicamento/Produto'].astype(str).str.lower()
    df_remume[col_remume] = df_remume[col_remume].astype(str).str.lower()

    for nome_remume in df_remume[col_remume]:
        correspondentes = df_estoque[df_estoque['Medicamento/Produto'].str.contains(nome_remume, na=False)]

        if not correspondentes.empty:
            melhor_correspondente = correspondentes['Medicamento/Produto'].iloc[0]
            total_estoque = correspondentes['Quantidade em Estoque'].sum()
        else:
            melhor_correspondente = ''
            total_estoque = 'EM FALTA'

        resultado.append({
            'Medicamento REMUME': nome_remume,
            'Correspondente Estoque': melhor_correspondente,
            'Quantidade em Estoque': total_estoque
        })

    df_resultado = pd.DataFrame(resultado)

    # Exportar para Excel
    hoje = datetime.today().strftime('%d-%m-%Y')
    nome_arquivo = f'Estoque_REMUME_atualizado_{hoje}.xlsx'
    caminho_saida = os.path.join('static', nome_arquivo)

    with pd.ExcelWriter(caminho_saida, engine='openpyxl') as writer:
        df_resultado.to_excel(writer, index=False, sheet_name='REMUME')
        ws = writer.sheets['REMUME']
        ws.insert_rows(1)
        ws.merge_cells('A1:C1')
        cell = ws['A1']
        cell.value = f'Estoque da REMUME por correspondência direta - {hoje}'
        cell.font = Font(bold=True)

    return nome_arquivo
