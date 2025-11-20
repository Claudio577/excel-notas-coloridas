import streamlit as st
import pandas as pd
import numpy as np
import tempfile
import re
from openpyxl import load_workbook
from openpyxl.styles import Font

st.title("📘 Unificador de Notas – 1º, 2º e 3º Bimestres (Notas Vermelhas < 5)")

st.write("""
Envie os 3 arquivos (1º, 2º e 3º bimestres).  
O sistema irá automaticamente:
- Limpar dados inúteis  
- Manter somente ALUNO + MATÉRIAS + NOTAS  
- Remover EP, ES, ET, AC  
- Remover colunas vazias  
- Unir por aluno  
- Pintar notas menores que 5 em vermelho  
""")

# ---------------------------------------------------------
# FUNÇÃO PARA LIMPAR CADA PLANILHA
# ---------------------------------------------------------

def limpar_planilha(file):
    df_raw = pd.read_excel(file, header=None)

    # Encontrar linha onde está escrito "ALUNO"
    linha_cabecalho = df_raw[df_raw.iloc[:, 0] == "ALUNO"].index[0]

    df = pd.read_excel(file, header=linha_cabecalho)

    # Manter somente alunos (nomes com 2 palavras ou mais)
    df = df[df["ALUNO"].str.contains(" ", na=False)]

    # Remover colunas Unnamed
    df = df.loc[:, ~df.columns.str.contains("Unnamed")]

    # Remover colunas inúteis
    df = df.drop(columns=["SITUAÇÃO", "TOTAL"], errors="ignore")

    # Extrair somente notas válidas
    def extrair_nota(valor):
        if pd.isna(valor):
            return np.nan
        nums = re.findall(r"\d+", str(valor))
        if not nums:
            return np.nan
        num = int(nums[0])
        return num if 0 <= num <= 10 else np.nan

    colunas_boas = ["ALUNO"]
    novas_colunas = {}

    for col in df.columns:
        if col == "ALUNO":
            continue

        df[col] = df[col].apply(extrair_nota)

        # Aceita somente colunas com pelo menos 1 nota
        if df[col].notna().sum() > 0:
            colunas_boas.append(col)
        else:
            continue

        # Limpar nome da matéria removendo códigos
        materia = re.split(r"\d+", col)[0].strip()
        if not materia:
            materia = col

        novas_colunas[col] = materia

    df = df[colunas_boas]
    df = df.rename(columns=novas_colunas)

    return df


# ---------------------------------------------------------
# UPLOAD DOS 3 BIMESTRES
# ---------------------------------------------------------

file_b1 = st.file_uploader("📤 Envie o Excel do 1º Bimestre", type=["xlsx"])
file_b2 = st.file_uploader("📤 Envie o Excel do 2º Bimestre", type=["xlsx"])
file_b3 = st.file_uploader("📤 Envie o Excel do 3º Bimestre", type=["xlsx"])

if file_b1 and file_b2 and file_b3:

    st.success("Arquivos carregados! Processando...")

    df1 = limpar_planilha(file_b1)
    df2 = limpar_planilha(file_b2)
    df3 = limpar_planilha(file_b3)

    # Renomear colunas com bimestre
    df1 = df1.rename(columns={c: f"{c}_B1" for c in df1.columns if c != "ALUNO"})
    df2 = df2.rename(columns={c: f"{c}_B2" for c in df2.columns if c != "ALUNO"})
    df3 = df3.rename(columns={c: f"{c}_B3" for c in df3.columns if c != "ALUNO"})

    # Unir todas por ALUNO
    df_final = df1.merge(df2, on="ALUNO", how="outer")
    df_final = df_final.merge(df3, on="ALUNO", how="outer")

    st.subheader("📄 Planilha Final (antes da coloração)")
    st.dataframe(df_final)

    # Salvar em arquivo temporário
    temp_out = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    df_final.to_excel(temp_out.name, index=False)

    # ---------------------------------------------------------
    # COLORIR NOTAS < 5 EM VERMELHO
    # ---------------------------------------------------------

    def colorir_notas(path):
        wb = load_workbook(path)
        ws = wb.active
        red_font = Font(color="FF0000", bold=True)

        # Ignorar coluna ALUNO
        for col in range(2, ws.max_column + 1):
            for row in range(2, ws.max_row + 1):
                cell = ws.cell(row=row, column=col)
                try:
                    if isinstance(cell.value, (int, float)) and cell.value < 5:
                        cell.font = red_font
                except:
                    pass

        wb.save(path)

    colorir_notas(temp_out.name)

    # ---------------------------------------------------------
    # BOTÃO DE DOWNLOAD
    # ---------------------------------------------------------

    with open(temp_out.name, "rb") as f:
        st.download_button(
            "⬇️ Baixar Planilha Final Unificada (Notas <5 em Vermelho)",
            f.read(),
            file_name="notas_bimestres_unificadas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

