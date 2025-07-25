import pandas as pd
from datetime import datetime, timedelta
import numpy as np
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import re

def normalizar_nome(med):
    if pd.isna(med):
        return ""

    med = med.lower()
    med = re.sub(r'\b(comprimido|cápsula|cápsulas|dr\u00e1gea|comprimidos|caps)\b', '', med)
    med = re.sub(r'\b(água para injeção|água destilada)\b', 'água', med)
    med = re.sub(r'[^a-zA-Z0-9\s]', '', med)
    med = re.sub(r'\s+', ' ', med).strip()
    return med

def executar_analise(df_estoque, df_dispensacao, df_distribuicao):
    df_estoque = df_estoque.dropna(subset=["Medicamento/Produto"])
    df_estoque["Medicamento/Produto Normalizado"] = df_estoque["Medicamento/Produto"].apply(normalizar_nome)

    df_dispensacao = df_dispensacao.dropna(subset=["Medicamento/Produto"])
    df_dispensacao["Medicamento/Produto Normalizado"] = df_dispensacao["Medicamento/Produto"].apply(normalizar_nome)

    df_distribuicao = df_distribuicao.dropna(subset=["Medicamento/Produto"])
    df_distribuicao["Medicamento/Produto Normalizado"] = df_distribuicao["Medicamento/Produto"].apply(normalizar_nome)

    df_dispensacao["Data"] = pd.to_datetime(df_dispensacao["Data Dispensação"], dayfirst=True, errors='coerce')
    df_distribuicao["Data"] = pd.to_datetime(df_distribuicao["Data Distribuição"], dayfirst=True, errors='coerce')

    hoje = datetime.today()
    trinta_dias_atras = hoje - timedelta(days=30)

    dispensado_30d = df_dispensacao[df_dispensacao["Data"] >= trinta_dias_atras]
    distribuido_30d = df_distribuicao[df_distribuicao["Data"] >= trinta_dias_atras]

    consumo_dispensado = dispensado_30d.groupby("Medicamento/Produto Normalizado")["Quantidade Dispensada"].sum()
    consumo_distribuido = distribuido_30d.groupby("Medicamento/Produto Normalizado")["Quantidade Distribuída"].sum()

    consumo_total = consumo_dispensado.add(consumo_distribuido, fill_value=0).rename("Total Consumido")

    df_estoque_group = df_estoque.groupby("Medicamento/Produto Normalizado").agg({
        "Quantidade em Estoque": "sum",
        "Validade": "first",
        "Medicamento/Produto": "first"
    }).reset_index()

    resultado = pd.merge(df_estoque_group, consumo_total, how="left", on="Medicamento/Produto Normalizado")
    resultado["Total Consumido"] = resultado["Total Consumido"].fillna(0)

    resultado["Consumo Mensal Médio Corrigido"] = resultado["Total Consumido"] / 1  # 30 dias = 1 mês

    resultado["Previsão Fim do Estoque"] = resultado.apply(lambda row: 
        (hoje + timedelta(days=(row["Quantidade em Estoque"] / row["Consumo Mensal Médio Corrigido"] * 30))
         if row["Consumo Mensal Médio Corrigido"] > 0 else pd.NaT), axis=1)

    resultado["Previsão Fim do Estoque"] = pd.to_datetime(resultado["Previsão Fim do Estoque"]).dt.date

    resultado["Análise de Dias com Falta"] = resultado.apply(lambda row: 
        f"FALTA DESDE {hoje.strftime('%d/%m/%Y')}" if row["Quantidade em Estoque"] == 0 else "", axis=1)

    colunas_finais = [
        "Medicamento/Produto", "Quantidade em Estoque", "Validade",
        "Total Consumido", "Consumo Mensal Médio Corrigido",
        "Previsão Fim do Estoque", "Análise de Dias com Falta"
    ]

    resultado_final = resultado[colunas_finais]

    caminho_saida = "resultado_analise.xlsx"
    resultado_final.to_excel(caminho_saida, index=False)

    # Destaque em vermelho para previsão de fim < 4 meses
    wb = load_workbook(caminho_saida)
    ws = wb.active

    red_fill = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")

    for row in range(2, ws.max_row + 1):
        data_previsao = ws[f"F{row}"].value
        if data_previsao and isinstance(data_previsao, datetime):
            if data_previsao.date() < (hoje + timedelta(days=120)).date():
                ws[f"F{row}"].fill = red_fill

    wb.save(caminho_saida)
    print(f"Análise salva em: {caminho_saida}")
