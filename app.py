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

    # Tem pelo menos duas palavras
    if len(partes) < 2:
        return False

    # Todas partes com letras
    if not all(p.isalpha() for p in partes):
        return False

    # Evitar EP, ES, ET, AC
    if len(partes[0]) <= 2:
        return False

    return True


# ---------------------------------------------------------------------
#  LIMPEZA DA PLANILHA
# ---------------------------------------------------------------------

def limpar_planilha(file):
    df_raw = pd.read_excel(file, header=None)

    # encontrar linha onde está "ALUNO"
    try:
        linha_cab = df_raw[df_raw.iloc[:, 0] == "ALUNO"].index[0]
    except:
        st.error("❌ A planilha enviada não contém a coluna 'ALUNO'.")
        st.stop()

    df = pd.read_excel(file, header=linha_cab)

    # manter apenas alunos válidos
    df = df[df["ALUNO"].apply(eh_aluno)]

    # remover colunas Unnamed
    df = df.loc[:, ~df.columns.str.contains("Unnamed")]

    # remover colunas inúteis
    df = df.drop(columns=["SITUAÇÃO", "TOTAL"], errors="ignore")

    # extrair nota numérica
    def extrair_nota(val):
        if pd.isna(val):
            return np.nan
        nums = re.findall(r"\d+", str(val))
        if not nums:
            return np.nan
        n = int(nums[0])
        return n if 0 <= n <= 10 else np.nan

    colunas_validas = ["ALUNO"]
    renomear = {}

    MATERIAS_REMOVER = ["Arte", "Esporte", "Música", "Artes", "Big A", "Inovação"]

    for col in df.columns:
        if col == "ALUNO":
            continue

        # limpeza da nota
        df[col] = df[col].apply(extrair_nota)

        # validação
        if df[col].notna().sum() == 0:
            continue

        # limpar nome da matéria
        materia = re.split(r"\d+", col)[0].strip()

        # remover matérias indesejadas
        if any(m.lower() in materia.lower() for m in MATERIAS_REMOVER):
            continue

        colunas_validas.append(col)
        renomear[col] = materia

    df = df[colunas_validas]
    df = df.rename(columns=renomear)
    return df


# ---------------------------------------------------------------------
# FORMATAR CABEÇALHO DUPLO SIMPLES (SEM MESCLAR)
# ---------------------------------------------------------------------

def formatar_cabecalho_simples(path, df_final):
    wb = load_workbook(path)
    ws = wb.active

    # adicionar 2 linhas no topo
    ws.insert_rows(1)
    ws.insert_rows(1)

    ws["A1"] = "ALUNO"
    ws["A2"] = ""

    col_excel = 2  # começa na coluna B

    for col in df_final.columns:
        if col == "ALUNO":
            continue

        materia, bi = col.split("_")
        bi_formatado = bi.replace("B", "ºBi")  # B1 → 1ºBi

        ws.cell(row=1, column=col_excel, value=materia)
        ws.cell(row=2, column=col_excel, value=bi_formatado)

        col_excel += 1

    wb.save(path)


# ---------------------------------------------------------------------
# COLORIR NOTAS < 5 EM VERMELHO
# ---------------------------------------------------------------------

def colorir_notas(path):
    wb = load_workbook(path)
    ws = wb.active
    red = Font(color="FF0000", bold=True)

    for col in range(2, ws.max_column + 1):
        for row in range(3, ws.max_row + 1):
            val = ws.cell(row=row, column=col).value
            try:
                if isinstance(val, (int, float)) and val < 5:
                    ws.cell(row=row, column=col).font = red
            except:
                pass

    wb.save(path)


# ---------------------------------------------------------------------
# UPLOAD
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

    # renomear colunas com bimestre
    df1 = df1.rename(columns={c: f"{c}_B1" for c in df1.columns if c != "ALUNO"})
    df2 = df2.rename(columns={c: f"{c}_B2" for c in df2.columns if c != "ALUNO"})
    df3 = df3.rename(columns={c: f"{c}_B3" for c in df3.columns if c != "ALUNO"})

    # unificar
    df_final = df1.merge(df2, on="ALUNO", how="outer")
    df_final = df_final.merge(df3, on="ALUNO", how="outer")

    df_final = df_final.fillna("–")

    # ordenar mantendo ordem do 1º bimestre
    df_final["ordem"] = df_final["ALUNO"].apply(
        lambda n: ordem_b1.index(n) if n in ordem_b1 else 999
    )
    df_final = df_final.sort_values("ordem").drop(columns=["ordem"])

    st.subheader("📄 Planilha Final (antes da formatação)")
    st.dataframe(df_final)

    # salvar arquivo
    temp_out = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    df_final.to_excel(temp_out.name, index=False)

    # aplicar formatações
    formatar_cabecalho_simples(temp_out.name, df_final)
    colorir_notas(temp_out.name)

    # botão download
    with open(temp_out.name, "rb") as f:
        st.download_button(
            "⬇️ Baixar Planilha Final Unificada",
            f.read(),
            file_name="notas_unificadas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
