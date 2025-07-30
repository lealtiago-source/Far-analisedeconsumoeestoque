import pandas as pd
from datetime import datetime
import re
import os
import unicodedata
from openpyxl.styles import Font
from openpyxl import load_workbook

def normalizar_nome(texto):
    if pd.isna(texto): return ''
    texto = str(texto).lower()
    texto = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('utf-8')
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

    df_estoque.columns = df_estoque.columns.str.strip()
    df_remume.columns = df_remume.columns.str.strip()

    nome_col_remume = 'RELAÇÃO MUNICIPAL DE MEDICAMENTOS ESSENCIAIS'

    df_estoque['normalizado'] = df_estoque['Medicamento/Produto'].apply(normalizar_nome)
    df_remume['normalizado'] = df_remume[nome_col_remume].apply(normalizar_nome)

    resultado = []

    for _, linha in df_remume.iterrows():
        nome_remume_original = linha[nome_col_remume]
        nome_remume_normalizado = linha['normalizado']

        # Filtra os correspondentes no estoque
        correspondentes = df_estoque[df_estoque['normalizado'] == nome_remume_normalizado]

        if not correspondentes.empty:
            linha_mais_recente = correspondentes.sort_values(by='Validade', ascending=False).iloc[0]
            nome_estoque = linha_mais_recente['Medicamento/Produto']
            quantidade = linha_mais_recente['Quantidade em Estoque']
        else:
            nome_estoque = ''
            quantidade = 'EM FALTA'

        resultado.append([nome_remume_original, nome_estoque, quantidade])

    df_resultado = pd.DataFrame(resultado, columns=[
        'Medicamento REMUME', 'Medicamento no Estoque', 'Quantidade em Estoque'
    ])

    hoje = datetime.today().strftime('%d-%m-%Y')
    nome_arquivo = f'Estoque_REMUME_atualizado_{hoje}.xlsx'
    caminho_saida = os.path.join('static', nome_arquivo)

    with pd.ExcelWriter(caminho_saida, engine='openpyxl') as writer:
        df_resultado.to_excel(writer, index=False, sheet_name='Correspondência')
        ws = writer.sheets['Correspondência']
        ws.insert_rows(1)
        ws.merge_cells('A1:C1')
        cell = ws['A1']
        cell.value = f'Correspondência Estoque vs REMUME - Gerado em {hoje}'
        cell.font = Font(bold=True)

    return nome_arquivo
