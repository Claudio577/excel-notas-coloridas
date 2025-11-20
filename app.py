import streamlit as st
import pandas as pd
import tempfile
import numpy as np
import re

st.title("📘 Extrator Inteligente de Notas – Limpeza Completa (v3)")

uploaded_file = st.file_uploader("Envie o Excel (.xlsx):", type=["xlsx"])

if uploaded_file:
    temp_input = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    temp_input.write(uploaded_file.getbuffer())
    temp_input.close()

    # Ler o arquivo sem cabeçalho
    df_raw = pd.read_excel(temp_input.name, header=None)

    # Buscar onde começa ALUNO
    linha_cabecalho = df_raw[df_raw.iloc[:, 0] == "ALUNO"].index[0]

    # Ler novamente com cabeçalho correto
    df = pd.read_excel(temp_input.name, header=linha_cabecalho)

    # Remover linhas sem aluno
    df = df.dropna(subset=["ALUNO"])

    # Remover colunas Unnamed
    df = df.loc[:, ~df.columns.str.contains("Unnamed")]

    # Remover colunas ruins
    df = df.drop(columns=["SITUAÇÃO", "TOTAL"], errors="ignore")

    # Função para extrair apenas números entre 0 e 10
    def extrair_nota(valor):
        if pd.isna(valor):
            return np.nan
        # encontrar números na célula
        nums = re.findall(r"\d+", str(valor))
        if not nums:
            return np.nan
        # converter para inteiro
        num = int(nums[0])
        # se número for nota válida
        if 0 <= num <= 10:
            return num
        return np.nan

    # Processar todas as colunas
    colunas_para_remover = []
    for col in df.columns:
        if col == "ALUNO":
            continue

        df[col] = df[col].apply(extrair_nota)

        # Se a coluna inteira ficou vazia → remover
        if df[col].dropna().empty:
            colunas_para_remover.append(col)

    # Remover colunas sem notas reais
    df = df.drop(columns=colunas_para_remover, errors="ignore")

    st.subheader("📄 Resultado Final – Apenas Notas Reais (0 a 10)")
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
