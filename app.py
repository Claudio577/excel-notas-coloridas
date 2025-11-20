import streamlit as st
import pandas as pd
import tempfile
import numpy as np

st.title("📘 Extrator de Notas – Alunos + Matérias + Notas Numéricas")

uploaded_file = st.file_uploader("Envie o Excel (.xlsx):", type=["xlsx"])

if uploaded_file:
    temp_input = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    temp_input.write(uploaded_file.getbuffer())
    temp_input.close()

    # Ler arquivo cru
    df_raw = pd.read_excel(temp_input.name, header=None)

    # Achar linha que contém "ALUNO"
    linha_cabecalho = df_raw[df_raw.iloc[:, 0] == "ALUNO"].index[0]

    # Ler com cabeçalho
    df = pd.read_excel(temp_input.name, header=linha_cabecalho)

    # Remover linhas sem nome de aluno
    df = df.dropna(subset=["ALUNO"])

    # Remover colunas Unnamed e SITUAÇÃO, TOTAL
    df = df.loc[:, ~df.columns.str.contains("Unnamed")]
    df = df.drop(columns=["SITUAÇÃO", "TOTAL"], errors="ignore")

    # Limpar todas as colunas numéricas:
    for col in df.columns:
        if col == "ALUNO":
            continue

        # Converter números; se não for número, vira NaN
        df[col] = pd.to_numeric(df[col], errors="coerce")

    st.subheader("📄 Resultado Final: Alunos + Todas as Matérias + Notas Numéricas")
    st.dataframe(df)

    # Salvar arquivo final
    temp_out = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    df.to_excel(temp_out.name, index=False)

    with open(temp_out.name, "rb") as f:
        st.download_button(
            "⬇️ Baixar Planilha Final",
            data=f.read(),
            file_name="notas_limpas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

