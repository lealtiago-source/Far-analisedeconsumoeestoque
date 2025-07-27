import pandas as pd
from datetime import datetime
import re
import os
import unidecode
from openpyxl.styles import Font
from openpyxl import load_workbook

def normalizar_nome(texto):
    if pd.isna(texto): return ''
    texto = unidecode.unidecode(str(texto).lower())
    texto = re.sub(r'\b(comprimido|capsula|solucao|injetavel|suspensao|dragea|frasco|ampola|ml|tablete|via oral|uso adulto|uso infantil)\b', '', texto)
    texto = re.sub(r'\b(cloreto|sodico|potassico|butilbrometo|maleato|nitrato|dihidrato|monoidratado|trihidratado|anidro)\b', '', texto)
    texto = re.sub(r'[^\w\s]', '', texto)
    texto = re.sub(r'\s+', ' ', texto).strip()

    match = re.search(r'([a-z\s]+)\s([\d]+(?:[.,]\d+)?\s*(mg|mcg|g|ml|mg/ml|%)?)', texto)
    if match:
        return f"{match.group(1).strip()} {match.group(2).strip()}"
    return texto

def executar_analise_remume(caminho_estoque, caminho_remume):
    df_estoque = pd.read_excel(caminho_estoque)
    df_remume = pd.read_excel(caminho_remume)

    # Nome correto da coluna no REMUME
    nome_col_remume = 'RELAÇÃO MUNICIPAL DE MEDICAMENTOS ESSENCIAIS'

    df_estoque['normalizado'] = df_estoque['Medicamento/Produto'].apply(normalizar_nome)
    df_remume['normalizado'] = df_remume[nome_col_remume].apply(normalizar_nome)

    # Medicamentos que estão na REMUME e no estoque
    intersecao = df_estoque[df_estoque['normalizado'].isin(df_remume['normalizado'])]

    # Medicamentos que estão na REMUME mas NÃO estão no estoque
    falta_no_estoque = df_remume[~df_remume['normalizado'].isin(df_estoque['normalizado'])].copy()
    falta_no_estoque['Quantidade em Estoque'] = 'EM FALTA'
    falta_no_estoque = falta_no_estoque[[nome_col_remume, 'Quantidade em Estoque']]
    falta_no_estoque = falta_no_estoque.rename(columns={nome_col_remume: 'Medicamento/Produto'})

    # Agrupar medicamentos em estoque (da REMUME)
    df_intersecao = intersecao.groupby('normalizado').agg({
        'Medicamento/Produto': 'last',
        'Quantidade em Estoque': 'sum'
    }).reset_index()

    df_intersecao = df_intersecao[['Medicamento/Produto', 'Quantidade em Estoque']]
    df_intersecao['tag'] = 'REMUME'
    falta_no_estoque['tag'] = 'REMUME'

    # Medicamentos que estão apenas no estoque
    apenas_estoque = df_estoque[~df_estoque['normalizado'].isin(df_remume['normalizado'])]
    df_apenas_estoque = apenas_estoque.groupby('normalizado').agg({
        'Medicamento/Produto': 'last',
        'Quantidade em Estoque': 'sum'
    }).reset_index()[['Medicamento/Produto', 'Quantidade em Estoque']]
    df_apenas_estoque['tag'] = 'OUTROS'

    # Concatenar: REMUME (encontrados + em falta), depois OUTROS
    df_final = pd.concat([df_intersecao, falta_no_estoque, df_apenas_estoque], ignore_index=True)

    # Ordenar por nome de medicamento
    df_final = df_final.sort_values(by=['tag', 'Medicamento/Produto'], key=lambda col: col.str.lower()).reset_index(drop=True)

    # Salvar Excel com título e data
    hoje = datetime.today().strftime('%d-%m-%Y')
    nome_arquivo = f'Estoque_REMUME_atualizado_{hoje}.xlsx'
    caminho_saida = os.path.join('static', nome_arquivo)

    with pd.ExcelWriter(caminho_saida, engine='openpyxl') as writer:
        df_final[['Medicamento/Produto', 'Quantidade em Estoque']].to_excel(writer, index=False, sheet_name='Estoque REMUME')
        ws = writer.sheets['Estoque REMUME']
        ws.insert_rows(1)
        ws.merge_cells('A1:B1')
        cell = ws['A1']
        cell.value = f'Estoque da REMUME - Gerado em {hoje}'
        cell.font = Font(bold=True)

    return nome_arquivo

