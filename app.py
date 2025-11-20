import streamlit as st
import pandas as pd
import numpy as np
import tempfile
import re
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, Side

st.title("📘 Unificador de Notas – 1º, 2º e 3º Bimestres (Formato de Boletim)")


# ---------------------------------------------------------------------
# FUNÇÃO PARA DETECTAR SE O TEXTO É UM NOME DE ALUNO REAL
# ---------------------------------------------------------------------
def eh_aluno(nome):
    if pd.isna(nome):
        return False
    partes = str(nome).split()
    if len(partes) < 2:
        return False
    if not all(p.isalpha() for p in partes):
        return False
    if len(partes[0]) <= 2:
        return False
    return True


# ---------------------------------------------------------------------
# LIMPEZA DA PLANILHA
# ---------------------------------------------------------------------
def limpar_planilha(file):
    df_raw = pd.read_excel(file, header=None)

    # Encontrar linha do cabeçalho (onde está "ALUNO")
    linha_cab = df_raw[df_raw.apply(lambda row: row.astype(str).str.contains("ALUNO").any(), axis=1)].index[0]
    df = pd.read_excel(file, header=linha_cab)

    # Manter apenas alunos reais
    df = df[df["ALUNO"].apply(eh_aluno)]

    # Remover colunas Unnamed
    df = df.loc[:, ~df.columns.str.contains("Unnamed")]

    # Remover colunas inúteis
    df = df.drop(columns=["SITUAÇÃO", "TOTAL"], errors="ignore")

    # Extrair apenas números das notas
    def extrair_nota(valor):
        if pd.isna(valor):
            return np.nan
        nums = re.findall(r"\d+", str(valor))
        if not nums:
            return np.nan
        n = int(nums[0])
        return n if 0 <= n <= 10 else np.nan

    colunas_validas = ["ALUNO"]
    renomear = {}

    for col in df.columns:
        if col == "ALUNO":
            continue

        df[col] = df[col].apply(extrair_nota)

        if df[col].notna().sum() > 0:
            colunas_validas.append(col)

            materia = re.split(r"\d+", col)[0].strip()
            if materia == "":
                materia = col

            renomear[col] = materia

    df = df[colunas_validas]
    df = df.rename(columns=renomear)

    return df


# LISTA FINAL DE MATÉRIAS (confirmada pelo usuário)
MATERIAS = [
    "Ciências",
    "Educação Financeira",
    "Educação Física",
    "Eletivas",
    "Geografia",
    "História",
    "Inglês",
    "Português",
    "Matemática",
    "Projeto de Vida",
    "Leitura",
    "Robótica"
]


# ---------------------------------------------------------------------
# UPLOAD DOS ARQUIVOS
# ---------------------------------------------------------------------
file_b1 = st.file_uploader("📤 Envie o Excel do 1º Bimestre", type=["xlsx"])
file_b2 = st.file_uploader("📤 Envie o Excel do 2º Bimestre", type=["xlsx"])
file_b3 = st.file_uploader("📤 Envie o Excel do 3º Bimestre", type=["xlsx"])

if file_b1 and file_b2 and file_b3:

    st.success("Arquivos carregados! Processando...")

    df1 = limpar_planilha(file_b1)
    df2 = limpar_planilha(file_b2)
    df3 = limpar_planilha(file_b3)

    # Renomear com _B1, _B2, _B3
    df1 = df1.rename(columns={c: f"{c}_B1" for c in df1.columns if c != "ALUNO"})
    df2 = df2.rename(columns={c: f"{c}_B2" for c in df2.columns if c != "ALUNO"})
    df3 = df3.rename(columns={c: f"{c}_B_3" for c in df3.columns if c != "ALUNO"})

    # Ordem original do B1
    ordem_b1 = df1["ALUNO"].tolist()

    # Unificar mantendo todos
    df_final = df1.merge(df2, on="ALUNO", how="outer")
    df_final = df_final.merge(df3, on="ALUNO", how="outer")

    # Preencher notas faltantes com "–"
    df_final = df_final.fillna("–")

    # Ordenar pela lista original
    df_final["ordem"] = df_final["ALUNO"].apply(
        lambda x: ordem_b1.index(x) if x in ordem_b1 else 999
    )
    df_final = df_final.sort_values("ordem").drop(columns="ordem")

    # Construir novo DataFrame no formato do boletim
    new_cols = ["ALUNO"]
    for materia in MATERIAS:
        for bi in ["B1", "B2", "B3"]:
            col_name = f"{materia}_{bi}"
            if col_name in df_final.columns:
                new_cols.append(col_name)
            else:
                df_final[col_name] = "–"
                new_cols.append(col_name)

    df_final = df_final[new_cols]

    st.subheader("📄 Planilha Final")
    st.dataframe(df_final, height=600)

    # -----------------------------------------------------------------
    #  EXPORTAR E ESTILIZAR COM DOIS CABEÇALHOS
    # -----------------------------------------------------------------
    temp_out = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    df_final.to_excel(temp_out.name, index=False, startrow=2)

    wb = load_workbook(temp_out.name)
    ws = wb.active

    # Criar cabeçalho duplo
    thin = Side(style="thin")
    border = Border(top=thin, left=thin, right=thin, bottom=thin)

    # Primeira linha: nomes das matérias
    col_index = 2  # Começa depois de ALUNO
    for materia in MATERIAS:
        ws.merge_cells(start_row=1, start_column=col_index, end_row=1, end_column=col_index + 2)
        ws.cell(row=1, column=col_index).value = materia
        ws.cell(row=1, column=col_index).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=1, column=col_index).border = border
        col_index += 3

    # Segunda linha: 1ºBi, 2ºBi, 3ºBi
    col_index = 2
    for materia in MATERIAS:
        ws.cell(row=2, column=col_index).value = "1ºBi"
        ws.cell(row=2, column=col_index+1).value = "2ºBi"
        ws.cell(row=2, column=col_index+2).value = "3ºBi"

        for c in range(col_index, col_index+3):
            ws.cell(row=2, column=c).alignment = Alignment(horizontal="center")
            ws.cell(row=2, column=c).border = border

        col_index += 3

    # Colorir notas < 5
    red = Font(color="FF0000", bold=True)
    for row in range(3, ws.max_row + 1):
        for col in range(2, ws.max_column + 1):
            val = ws.cell(row=row, column=col).value
            if isinstance(val, (int, float)) and val < 5:
                ws.cell(row=row, column=col).font = red

    wb.save(temp_out.name)

    # Botão de download
    with open(temp_out.name, "rb") as f:
        st.download_button(
            "⬇️ Baixar Planilha Final Estilizada",
            f.read(),
            file_name="notas_formatadas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

