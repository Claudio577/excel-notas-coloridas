import streamlit as st
import pandas as pd
import tempfile
import numpy as np
import re

st.title("📘 Extrator Inteligente de Notas – Limpeza Completa (v5)")

uploaded_file = st.file_uploader("Envie o Excel (.xlsx):", type=["xlsx"])

if uploaded_file:
    temp_input = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    temp_input.write(uploaded_file.getbuffer())
    temp_input.close()

    df_raw = pd.read_excel(temp_input.name, header=None)

    linha_cabecalho = df_raw[df_raw.iloc[:, 0] == "ALUNO"].index[0]

    df = pd.read_excel(temp_input.name, header=linha_cabecalho)

    df = df.dropna(subset=["ALUNO"])
    df = df.loc[:, ~df.columns.str.contains("Unnamed")]
    df = df.drop(columns=["SITUAÇÃO", "TOTAL"], errors="ignore")

    def extrair_nota(valor):
        if pd.isna(valor):
            return np.nan
        nums = re.findall(r"\d+", str(valor))
        if not nums:
            return np.nan
        num = int(nums[0])
        if 0 <= num <= 10:
            return num
        return np.nan

    colunas_para_remover = []
    for col in df.columns:
        if col == "ALUNO":
            continue
        df[col] = df[col].apply(extrair_nota)
        if df[col].dropna().empty:
            colunas_para_remover.append(col)

    df = df.drop(columns=colunas_para_remover, errors="ignore")

    def limpar_nome_coluna(nome):
        base = re.split(r"\d+", nome)[0].strip()
        return base if base else nome

    df.columns = [limpar_nome_coluna(col) for col in df.columns]

    st.subheader("📄 Resultado Final – Matérias Renomeadas")
    st.dataframe(df)

    temp_out = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    df.to_excel(temp_out.name, index=False)

    with open(temp_out.name, "rb") as f:
        st.download_button(
            "⬇️ Baixar Planilha Final Completa",
            data=f.read(),
            file_name="notas_limpas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # ---- BOTÃO PARA NOTAS VERMELHAS ----
    st.subheader("📌 Filtrar alunos com notas vermelhas (<5)")

    if st.button("Mostrar somente alunos com nota vermelha"):
        col_notas = [c for c in df.columns if c != "ALUNO"]
        filtro = df[col_notas].lt(5).any(axis=1)
        df_vermelhas = df[filtro]

        st.subheader("🚨 Alunos com notas vermelhas:")
        st.dataframe(df_vermelhas)

        temp_vermelhas = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        df_vermelhas.to_excel(temp_vermelhas.name, index=False)

        with open(temp_vermelhas.name, "rb") as f:
            st.download_button(
                "⬇️ Baixar alunos com notas vermelhas",
                data=f.read(),
                file_name="notas_vermelhas.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
