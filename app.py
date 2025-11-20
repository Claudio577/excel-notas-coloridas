import streamlit as st
import pandas as pd
import numpy as np
import tempfile
import re
from openpyxl import load_workbook
from openpyxl.styles import Font

st.title("📘 Unificador de Notas – 1º, 2º e 3º Bimestres")

st.write("""
Envie os 3 arquivos (um por bimestre). O sistema irá:
- Limpar os dados
- Remover colunas inúteis
- Extrair somente notas (0–10)
- Renomear matérias automaticamente
- Juntar tudo em uma única planilha organizada por bimestre
- Pintar notas < 5 em vermelho
""")

# ------------------ FUNÇÃO DE LIMPEZA ------------------

def limpar_planilha(file):
    df_raw = pd.read_excel(file, header=None)

    # Encontrar a linha do cabeçalho (onde tem "ALUNO")
    linha_cabecalho = df_raw[df_raw.iloc[:, 0] == "ALUNO"].index[0]

    # Ler com cabeçalho correto
    df = pd.read_excel(file, header=linha_cabecalho)

    # Remover linhas sem aluno
    df = df.dropna(subset=["ALUNO"])

    # Remover colunas Unnamed e colunas inúteis
    df = df.loc[:, ~df.columns.str.contains("Unnamed")]
    df = df.drop(columns=["SITUAÇÃO", "TOTAL"], errors="ignore")

    # Função para extrair apenas números de 0 a 10
    def extrair_nota(valor):
        if pd.isna(valor):
            return np.nan
        nums = re.findall(r"\d+", str(valor))
        if not nums:
            return np.nan
        num = int(nums[0])
        if 0 <= num <= 10:
            return num
        return np.nan

    # Limpar todas as colunas
    colunas_para_remover = []
    for col in df.columns:
        if col == "ALUNO":
            continue
        df[col] = df[col].apply(extrair_nota)

        # Remove coluna sem nenhuma nota válida
        if df[col].dropna().empty:
            colunas_para_remover.append(col)

    df = df.drop(columns=colunas_para_remover, errors="ignore")

    # Renomear coluna: remover códigos numéricos
    def limpar_nome_coluna(nome):
        base = re.split(r"\d+", nome)[0].strip()
        return base if base else nome

    df.columns = [limpar_nome_coluna(col) for col in df.columns]

    return df


# ------------------ UPLOAD DOS 3 BIMESTRES ------------------

file_b1 = st.file_uploader("Envie o Excel do 1º Bimestre", type=["xlsx"])
file_b2 = st.file_uploader("Envie o Excel do 2º Bimestre", type=["xlsx"])
file_b3 = st.file_uploader("Envie o Excel do 3º Bimestre", type=["xlsx"])

if file_b1 and file_b2 and file_b3:

    st.success("Arquivos carregados! Processando...")

    df1 = limpar_planilha(file_b1)
    df2 = limpar_planilha(file_b2)
    df3 = limpar_planilha(file_b3)

    st.write("🔍 Prévia 1º Bimestre:")
    st.dataframe(df1.head())

    # Renomear colunas para incluir B1, B2 e B3
    df1 = df1.rename(columns={col: f"{col}_B1" for col in df1.columns if col != "ALUNO"})
    df2 = df2.rename(columns={col: f"{col}_B2" for col in df2.columns if col != "ALUNO"})
    df3 = df3.rename(columns={col: f"{col}_B3" for col in df3.columns if col != "ALUNO"})

    # Unir todas pelo nome do aluno
    df_final = df1.merge(df2, on="ALUNO", how="outer")
    df_final = df_final.merge(df3, on="ALUNO", how="outer")

    st.subheader("📄 Planilha Final (Sem Coloração Ainda)")
    st.dataframe(df_final)

    # ------------------ GERAR EXCEL COM COLORAÇÃO ------------------

    temp_out = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    df_final.to_excel(temp_out.name, index=False)

    def colorir_notas(caminho):
        wb = load_workbook(caminho)
        ws = wb.active

        red_font = Font(color="FF0000", bold=True)

        for col in range(2, ws.max_column + 1):
            for row in range(2, ws.max_row + 1):
                cell = ws.cell(row=row, column=col)

                try:
                    if isinstance(cell.value, (int, float)) and cell.value < 5:
                        cell.font = red_font
                except:
                    pass

        wb.save(caminho)

    colorir_notas(temp_out.name)

    with open(temp_out.name, "rb") as f:
        st.download_button(
            "⬇️ Baixar Planilha Final com B1 + B2 + B3 (Notas < 5 em Vermelho)",
            data=f.read(),
            file_name="notas_bimestres_unificadas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
