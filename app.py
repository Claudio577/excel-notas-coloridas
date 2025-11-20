import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import openpyxl
from openpyxl.styles import PatternFill

# ----------------------------------------------------------
#  FUNÇÃO DEFINITIVA DE LIMPEZA — À PROVA DE QUALQUER MAPÃO
# ----------------------------------------------------------
def limpar_planilha(df):

    # 1) Normaliza QUALQUER conteúdo da célula → texto simples
    def normalizar_celula(x):
        if isinstance(x, (list, tuple, set)):
            return " ".join(map(str, x))

        if isinstance(x, pd.Series):
            return " ".join(map(str, x.tolist()))

        if isinstance(x, pd.DataFrame):
            return " ".join(map(str, x.values.flatten().tolist()))

        if isinstance(x, np.ndarray):
            return " ".join(map(str, x.tolist()))

        return "" if pd.isna(x) else str(x).strip()

    df = df.applymap(normalizar_celula)
    df = df.replace(["nan", "None"], "")

    # Remove colunas e linhas vazias
    df = df.loc[:, (df != "").any(axis=0)]
    df = df[(df != "").any(axis=1)]

    # 2) Encontrar linha do cabeçalho (linha com "ALUNO")
    linha_aluno = None
    for i, row in df.iterrows():
        if row.astype(str).str.contains("ALUNO", case=False).any():
            linha_aluno = i
            break

    if linha_aluno is None:
        st.error("❌ ERRO: Não encontrei a linha que contém 'ALUNO'.")
        return None

    df.columns = df.iloc[linha_aluno]
    df = df.iloc[linha_aluno + 1:]

    # Remove colunas Unnamed
    df = df.loc[:, ~df.columns.astype(str).str.contains("Unnamed")]

    # 3) Manter APENAS colunas numéricas + ALUNO
    colunas_ok = ["ALUNO"]

    for col in df.columns:
        if col == "ALUNO":
            continue

        test = pd.to_numeric(df[col], errors="coerce")

        if test.notna().sum() > 0:
            colunas_ok.append(col)

    df = df[colunas_ok]

    # Converte notas para números
    for col in df.columns:
        if col != "ALUNO":
            df[col] = pd.to_numeric(df[col], errors="coerce")

    return df


# ----------------------------------------------------------
#  DESTACAR NOTAS MENORES QUE 5 (Vermelho)
# ----------------------------------------------------------
def aplicar_cor_vermelha(df):
    vermelho = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

    wb = openpyxl.Workbook()
    ws = wb.active

    # Cabeçalho
    ws.append(df.columns.tolist())

    # Linhas
    for _, row in df.iterrows():
        ws.append(row.tolist())

    # Aplicar cor
    for row in ws.iter_rows(min_row=2):
        for cell in row[1:]:
            try:
                if float(cell.value) < 5:
                    cell.fill = vermelho
            except:
                pass

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio


# ----------------------------------------------------------
#  INTERFACE STREAMLIT
# ----------------------------------------------------------
st.title("🟦 Unificador de Notas — 1º, 2º e 3º Bimestre (com notas vermelhas)")

st.write("Envie **3 planilhas** do MAPÃO (1º, 2º e 3º bimestres).")

uploaded_b1 = st.file_uploader("📘 1º Bimestre", type=["xlsx"])
uploaded_b2 = st.file_uploader("📙 2º Bimestre", type=["xlsx"])
uploaded_b3 = st.file_uploader("📗 3º Bimestre", type=["xlsx"])

if uploaded_b1 and uploaded_b2 and uploaded_b3:
    st.success("✔ Arquivos carregados! Processando...")

    # Processar cada bimestre
    df1 = limpar_planilha(pd.read_excel(uploaded_b1, header=None))
    df2 = limpar_planilha(pd.read_excel(uploaded_b2, header=None))
    df3 = limpar_planilha(pd.read_excel(uploaded_b3, header=None))

    if df1 is not None and df2 is not None and df3 is not None:

        # Renomear colunas para indicar bimestre
        df1 = df1.add_suffix("_B1")
        df2 = df2.add_suffix("_B2")
        df3 = df3.add_suffix("_B3")

        df1 = df1.rename(columns={"ALUNO_B1": "ALUNO"})
        df2 = df2.rename(columns={"ALUNO_B2": "ALUNO"})
        df3 = df3.rename(columns={"ALUNO_B3": "ALUNO"})

        # Unificar
        df_final = df1.merge(df2, on="ALUNO", how="outer")
        df_final = df_final.merge(df3, on="ALUNO", how="outer")

        st.subheader("📄 Planilha Final (antes da coloração)")
        st.dataframe(df_final)

        # Criar Excel colorido
        excel_final = aplicar_cor_vermelha(df_final)

        st.download_button(
            label="⬇ Baixar Planilha Final (notas vermelhas <5)",
            data=excel_final,
            file_name="notas_bimestres_unificadas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

else:
    st.info("Envie as **3 planilhas** para continuar.")
