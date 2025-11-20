import streamlit as st
import pandas as pd
import numpy as np
import tempfile
import re
from openpyxl import load_workbook
from openpyxl.styles import Font


st.title("📘 Unificador de Notas – 1º, 2º e 3º Bimestres (Notas Vermelhas < 5)")


# ------------------------------------------------------------
# Acha linha e coluna onde está "ALUNO"
# ------------------------------------------------------------
def encontrar_linha_e_coluna_aluno(df):
    for i in range(df.shape[0]):
        for j in range(df.shape[1]):
            if str(df.iat[i, j]).strip().upper() == "ALUNO":
                return i, j
    return None, None


# ------------------------------------------------------------
# Detecta se é nome real de aluno
# ------------------------------------------------------------
def eh_aluno(nome):
    if pd.isna(nome):
        return False
    partes = nome.split()
    if len(partes) < 2:
        return False
    if not all(p.isalpha() for p in partes):
        return False
    if len(partes[0]) <= 2:
        return False
    return True


# ------------------------------------------------------------
# Limpa planilha individual
# ------------------------------------------------------------
def limpar_planilha(file, sufixo):

    df_raw = pd.read_excel(file, header=None)

    linha, coluna = encontrar_linha_e_coluna_aluno(df_raw)
    if linha is None:
        raise ValueError("Erro: coluna ALUNO não encontrada no Excel.")

    df = pd.read_excel(file, header=linha)

    # Remover linhas que não são alunos
    df = df[df["ALUNO"].apply(eh_aluno)]

    # Remover colunas sem nome ou Unnamed
    df = df.loc[:, ~df.columns.str.contains("Unnamed")]

    # Remover colunas inúteis se existirem
    df = df.drop(columns=["SITUAÇÃO", "TOTAL"], errors="ignore")

    # Extrair números (notas)
    def extrair_nota(v):
        if pd.isna(v):
            return np.nan
        nums = re.findall(r"\d+", str(v))
        if not nums:
            return np.nan
        val = int(nums[0])
        return val if 0 <= val <= 10 else np.nan

    colunas_ok = ["ALUNO"]
    renomear_cols = {}

    for col in df.columns:
        if col == "ALUNO":
            continue

        df[col] = df[col].apply(extrair_nota)

        if df[col].notna().sum() == 0:
            continue

        colunas_ok.append(col)

        materia = re.split(r"\d+", col)[0].strip()
        if materia == "":
            materia = col

        renomear_cols[col] = f"{materia}_{sufixo}"

    df = df[colunas_ok]
    df = df.rename(columns=renomear_cols)

    return df.reset_index(drop=True)


# ------------------------------------------------------------
# Agrupar matérias: MAT_B1, MAT_B2, MAT_B3 lado a lado
# ------------------------------------------------------------
def organizar_por_materia(df):
    colunas = ["ALUNO"]

    materias = sorted({c.rsplit("_", 1)[0] for c in df.columns if c != "ALUNO"})

    for materia in materias:
        for bim in ["B1", "B2", "B3"]:
            nome_col = f"{materia}_{bim}"
            if nome_col in df.columns:
                colunas.append(nome_col)

    return df[colunas]


# ------------------------------------------------------------
# Colorir notas vermelhas
# ------------------------------------------------------------
def colorir_notas(path):
    wb = load_workbook(path)
    ws = wb.active
    red = Font(color="FF0000", bold=True)

    for col in range(2, ws.max_column + 1):
        for row in range(2, ws.max_row + 1):
            cell = ws.cell(row=row, column=col)
            try:
                if isinstance(cell.value, (int, float)) and cell.value < 5:
                    cell.font = red
            except:
                pass

    wb.save(path)


# ------------------------------------------------------------
# Uploads
# ------------------------------------------------------------
file_b1 = st.file_uploader("📤 Envie o Excel do 1º Bimestre", type=["xlsx"])
file_b2 = st.file_uploader("📤 Envie o Excel do 2º Bimestre", type=["xlsx"])
file_b3 = st.file_uploader("📤 Envie o Excel do 3º Bimestre", type=["xlsx"])


if file_b1 and file_b2 and file_b3:

    st.success("✔ Arquivos carregados!")

    df1 = limpar_planilha(file_b1, "B1")
    df2 = limpar_planilha(file_b2, "B2")
    df3 = limpar_planilha(file_b3, "B3")

    df_final = df1.merge(df2, on="ALUNO", how="outer")
    df_final = df_final.merge(df3, on="ALUNO", how="outer")

    # AGRUPAR AS MATÉRIAS
    df_final = organizar_por_materia(df_final)

    st.subheader("📄 Planilha Final Organizada (antes da cor)")
    st.dataframe(df_final, height=500)

    # salvar temporário
    temp_out = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    df_final.to_excel(temp_out.name, index=False)

    # aplicar cor
    colorir_notas(temp_out.name)

    with open(temp_out.name, "rb") as f:
        st.download_button(
            "⬇️ Baixar Planilha Final Agrupada e Colorida",
            f.read(),
            file_name="notas_unificadas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
