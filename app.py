import streamlit as st
import pandas as pd
import numpy as np
import tempfile
import re
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

st.title("📘 Unificador de Notas – Cabeçalho Agrupado por Matéria")


# ---------------------------------------------------------
# FUNÇÃO PARA DETECTAR ALUNO
# ---------------------------------------------------------
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


# ---------------------------------------------------------
# LIMPEZA DE PLANILHA
# ---------------------------------------------------------
def limpar_planilha(file):
    df_raw = pd.read_excel(file, header=None)

    try:
        linha_cab = df_raw[df_raw.iloc[:, 0] == "ALUNO"].index[0]
    except:
        st.error("❌ Planilha sem coluna 'ALUNO'.")
        st.stop()

    df = pd.read_excel(file, header=linha_cab)

    df = df[df["ALUNO"].apply(eh_aluno)]
    df = df.loc[:, ~df.columns.str.contains("Unnamed")]
    df = df.drop(columns=["SITUAÇÃO", "TOTAL"], errors="ignore")

    def extrair_nota(v):
        if pd.isna(v):
            return np.nan
        nums = re.findall(r"\d+", str(v))
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

        if df[col].notna().sum() == 0:
            continue

        colunas_validas.append(col)
        materia = re.split(r"\d+", col)[0].strip()
        if materia == "":
            materia = col
        renomear[col] = materia

    df = df[colunas_validas]
    df = df.rename(columns=renomear)

    return df


# ---------------------------------------------------------
# UPLOAD
# ---------------------------------------------------------
file_b1 = st.file_uploader("📤 1º Bimestre", type=["xlsx"])
file_b2 = st.file_uploader("📤 2º Bimestre", type=["xlsx"])
file_b3 = st.file_uploader("📤 3º Bimestre", type=["xlsx"])

if file_b1 and file_b2 and file_b3:

    st.success("Arquivos carregados! Processando...")

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
        lambda n: ordem_b1.index(n) if n in ordem_b1 else 999
    )

    df_final = df_final.sort_values("ordem").drop(columns=["ordem"])

    # Salvar temporário para formatação
    temp_out = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    df_final.to_excel(temp_out.name, index=False, header=False)

    # ---------------------------------------------------------
    # APLICAR CABEÇALHO MESCLADO
    # ---------------------------------------------------------
    wb = load_workbook(temp_out.name)
    ws = wb.active

    colunas = df_final.columns.tolist()

    materias = {}
    for idx, col in enumerate(colunas):
        if col == "ALUNO":
            continue
        materia, bim = col.split("_")
        bim = bim.replace("B", "") + "ºBi"
        if materia not in materias:
            materias[materia] = []
        materias[materia].append((idx + 1, bim))  # coluna Excel inicia em 1

    # Criar cabeçalho com mesclagem
    ws.insert_rows(1)
    ws.insert_rows(1)

    ws["A1"] = "ALUNO"
    ws["A2"] = ""

    for materia, cols in materias.items():
        col_inicio = cols[0][0] + 1
        col_fim = cols[-1][0] + 1

        ws.merge_cells(start_row=1, start_column=col_inicio,
                       end_row=1, end_column=col_fim)
        ws.cell(row=1, column=col_inicio).value = materia
        ws.cell(row=1, column=col_inicio).alignment = Alignment(horizontal="center")

        for i, (_, nome_bi) in enumerate(cols):
            ws.cell(row=2, column=col_inicio + i).value = nome_bi
            ws.cell(row=2, column=col_inicio + i).alignment = Alignment(horizontal="center")

    # ---------------------------------------------------------
    # COLORIR NOTAS < 5
    # ---------------------------------------------------------
    red = Font(color="FF0000", bold=True)

    for r in range(3, ws.max_row + 1):
        for c in range(2, ws.max_column + 1):
            val = ws.cell(r, c).value
            try:
                if isinstance(val, (int, float)) and val < 5:
                    ws.cell(r, c).font = red
            except:
                pass

    wb.save(temp_out.name)

    # ---------------------------------------------------------
    # DOWNLOAD
    # ---------------------------------------------------------
    with open(temp_out.name, "rb") as f:
        st.download_button(
            "⬇️ Baixar Planilha Formatada",
            f.read(),
            file_name="notas_unificadas_formatadas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

