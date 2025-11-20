import streamlit as st
import pandas as pd
import numpy as np
import tempfile
import re
from openpyxl import load_workbook
from openpyxl.styles import Font

st.title("📘 Unificador de Notas – 1º, 2º e 3º Bimestres (Notas Vermelhas < 5)")


# ---------------------------------------------------------------------
#  FUNÇÃO PARA DETECTAR SE O TEXTO É UM NOME DE ALUNO REAL
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
#  LIMPEZA DA PLANILHA
# ---------------------------------------------------------------------
def limpar_planilha(file):
    df_raw = pd.read_excel(file, header=None)

    try:
        linha_cab = df_raw[df_raw.iloc[:, 0] == "ALUNO"].index[0]
    except:
        st.error("❌ A planilha enviada não contém uma coluna 'ALUNO'.")
        st.stop()

    df = pd.read_excel(file, header=linha_cab)

    df = df[df["ALUNO"].apply(eh_aluno)]

    df = df.loc[:, ~df.columns.str.contains("Unnamed")]

    df = df.drop(columns=["SITUAÇÃO", "TOTAL"], errors="ignore")

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

    # Remover matérias proibidas
    materias_excluir = ["ARTE", "Arte", "AR", "Esporte", "Música", "Big", "Inovação"]

    for col in df.columns:
        if col == "ALUNO":
            continue

        if any(m in col for m in materias_excluir):
            continue

        df[col] = df[col].apply(extrair_nota)

        if df[col].notna().sum() > 0:
            colunas_validas.append(col)
        else:
            continue

        materia = re.split(r"\d+", col)[0].strip()
        if materia == "":
            materia = col

        renomear[col] = materia

    df = df[colunas_validas]
    df = df.rename(columns=renomear)

    return df


# ---------------------------------------------------------------------
#  UPLOAD DOS 3 BIMESTRES
# ---------------------------------------------------------------------
file_b1 = st.file_uploader("📤 Envie o Excel do 1º Bimestre", type=["xlsx"])
file_b2 = st.file_uploader("📤 Envie o Excel do 2º Bimestre", type=["xlsx"])
file_b3 = st.file_uploader("📤 Envie o Excel do 3º Bimestre", type=["xlsx"])

if file_b1 and file_b2 and file_b3:

    st.success("Arquivos carregados! Limpando dados...")

    df1 = limpar_planilha(file_b1)
    df2 = limpar_planilha(file_b2)
    df3 = limpar_planilha(file_b3)

    ordem_b1 = df1["ALUNO"].tolist()

    df1 = df1.rename(columns={c: f"{c}_B1" for c in df1.columns if c != "ALUNO"})
    df2 = df2.rename(columns={c: f"{c}_B2" for c in df2.columns if c != "ALUNO"})
    df3 = df3.rename(columns={c: f"{c}_B3" for c in df3.columns if c != "ALUNO"})

    df_final = df1.merge(df2, on="ALUNO", how="outer")
    df_final = df_final.merge(df3, on="ALUNO", how="outer")

    df_final = df_final.fillna("–")

    df_final["ordem"] = df_final["ALUNO"].apply(
        lambda nome: ordem_b1.index(nome) if nome in ordem_b1 else 999
    )
    df_final = df_final.sort_values("ordem").drop(columns=["ordem"])

    st.subheader("📄 Planilha Final (antes da formatação)")
    st.dataframe(df_final)


    # ---------------------------------------------------------------------
    #  SALVAR + GERAR CABEÇALHO DUPLO + COLORIR < 5
    # ---------------------------------------------------------------------
    temp_out = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    df_final.to_excel(temp_out.name, index=False)

    # ---------------------------------------------------------
    #  CABEÇALHO DUPLO
    # ---------------------------------------------------------
    def formatar_cabecalho_duplo(path):
        wb = load_workbook(path)
        ws = wb.active

        materias = {}
        for col in df_final.columns:
            if col == "ALUNO":
                continue
            nome, bi = col.split("_")  # EX: Ciências_B1
            if nome not in materias:
                materias[nome] = []
            materias[nome].append(col)

        ws.insert_rows(1)
        ws.insert_rows(1)

        ws["A1"] = ""
        ws["A2"] = "ALUNO"

        col_excel = 2

        for materia, colunas in materias.items():
            n_cols = len(colunas)

            c1 = ws.cell(row=1, column=col_excel)
            c1.value = materia
            ws.merge_cells(start_row=1, start_column=col_excel,
                           end_row=1, end_column=col_excel + n_cols - 1)

            for i, colname in enumerate(colunas):
                bi_raw = colname.split("_")[1]  # B1,B2,B3
                bi = bi_raw.replace("B", "ºBi")
                ws.cell(row=2, column=col_excel + i, value=bi)

            col_excel += n_cols

        wb.save(path)

    # ---------------------------------------------------------
    # COLORIR NOTAS < 5
    # ---------------------------------------------------------
    def colorir_notas(path):
        wb = load_workbook(path)
        ws = wb.active
        red = Font(color="FF0000", bold=True)

        for col in range(2, ws.max_column + 1):
            for row in range(3, ws.max_row + 1):  # agora notas começam na linha 3
                val = ws.cell(row=row, column=col).value
                if isinstance(val, (int, float)) and val < 5:
                    ws.cell(row=row, column=col).font = red

        wb.save(path)

    formatar_cabecalho_duplo(temp_out.name)
    colorir_notas(temp_out.name)

    # ---------------------------------------------------------------------
    # DOWNLOAD
    # ---------------------------------------------------------------------
    with open(temp_out.name, "rb") as f:
        st.download_button(
            "⬇️ Baixar Planilha Final Unificada FORMATADA",
            f.read(),
            file_name="notas_unificadas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
