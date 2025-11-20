import streamlit as st
import pandas as pd
import tempfile
import numpy as np
import re

st.title("📘 Extrator Inteligente de Notas – Limpeza Automática")

uploaded_file = st.file_uploader("Envie o Excel (.xlsx):", type=["xlsx"])

if uploaded_file:
    temp_input = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    temp_input.write(uploaded_file.getbuffer())
    temp_input.close()

    # Ler arquivo cru
    df_raw = pd.read_excel(temp_input.name, header=None)

    # Achar linha de cabeçalho (onde começa ALUNO)
    linha_cabecalho = df_raw[df_raw.iloc[:, 0] == "ALUNO"].index[0]

    # Ler com cabeçalho correto
    df = pd.read_excel(temp_input.name, header=linha_cabecalho)

    # Remover linhas vazias e colunas Unnamed
    df = df.dropna(subset=["ALUNO"])
    df = df.loc[:, ~df.columns.str.contains("Unnamed")]

    # Remover colunas desnecessárias
    df = df.drop(columns=["SITUAÇÃO", "TOTAL"], errors="ignore")

    # Processar cada coluna
    colunas_para_remover = []

    for col in df.columns:
        if col == "ALUNO":
            continue

        # Extrair números usando regex: pegamos somente o primeiro número da célula
        df[col] = df[col].astype(str).apply(lambda x: re.findall(r"\d+", x))
        df[col] = df[col].apply(lambda x: int(x[0]) if x else np.nan)

        # Se a coluna não possuir nenhum número → remover
        if df[col].dropna().empty:
            colunas_para_remover.append(col)

    # Remover colunas sem números (ex.: MÚSICA, ARTE com letras)
    df = df.drop(columns=colunas_para_remover, errors="ignore")

    st.subheader("📄 Resultado Final – Colunas Limpas e Corrigidas")
    st.dataframe(df)

    # Salvar Excel final
    temp_out = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    df.to_excel(temp_out.name, index=False)

    with open(temp_out.name, "rb") as f:
        st.download_button(
            "⬇️ Baixar Planilha Final",
            data=f.read(),
            file_name="notas_limpas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
