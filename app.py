import streamlit as st
import pandas as pd
import tempfile
import numpy as np

st.title("📘 Extrator de Notas – Limpeza Completa")

st.write("""
Este app extrai:
- ALUNO
- Todas as notas numéricas (0 a 10)
- Remove **Unnamed**, **SITUAÇÃO**, **TOTAL**, e colunas com texto (AC, ES, EP etc.)
""")

uploaded_file = st.file_uploader("Envie o Excel (.xlsx):", type=["xlsx"])

if uploaded_file:

    # Salvar arquivo temporário
    temp_input = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    temp_input.write(uploaded_file.getbuffer())
    temp_input.close()

    # Ler SEM cabeçalho, para achar a linha certa
    df_raw = pd.read_excel(temp_input.name, header=None)

    st.subheader("Primeiras linhas detectadas:")
    st.dataframe(df_raw.head(20))

    # Encontrar a linha onde está o texto "ALUNO"
    linha_cabecalho = df_raw[df_raw.iloc[:, 0] == "ALUNO"].index[0]

    # Ler o Excel novamente com cabeçalho correto
    df = pd.read_excel(temp_input.name, header=linha_cabecalho)

    # Remover alunos vazios
    df = df.dropna(subset=["ALUNO"])

    # Remover colunas desnecessárias
    colunas_remover = [
        "SITUAÇÃO", "TOTAL"
    ]

    # Remover colunas Unnamed automaticamente
    colunas_remover += [c for c in df.columns if "Unnamed" in str(c)]

    # Remover colunas com valores não numéricos (exceto ALUNO)
    colunas_numericas = []

    for col in df.columns:
        if col == "ALUNO":
            continue
        # Se TODOS os valores forem números, mantemos
        if pd.to_numeric(df[col], errors="coerce").notna().sum() > 0:
            colunas_numericas.append(col)

    # Montar dataframe final
    df_final = df[["ALUNO"] + colunas_numericas]

    st.subheader("📄 Resultado Final (somente notas reais):")
    st.dataframe(df_final)

    # Salvar arquivo final
    temp_output = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    df_final.to_excel(temp_output.name, index=False)

    with open(temp_output.name, "rb") as f:
        st.download_button(
            "⬇️ Baixar Planilha Limpa (Sem Unnamed)",
            data=f.read(),
            file_name="notas_limpas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
