import streamlit as st
import pandas as pd
import tempfile

st.title("📘 Extrator de Notas – Somente Notas Numéricas")

uploaded_file = st.file_uploader("Envie seu Excel (.xlsx):", type=["xlsx"])

if uploaded_file:
    temp_input = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    temp_input.write(uploaded_file.getbuffer())
    temp_input.close()

    df_raw = pd.read_excel(temp_input.name, header=None)

    # Encontrar linha do cabeçalho
    linha_cabecalho = df_raw[df_raw.iloc[:, 0] == "ALUNO"].index[0]

    # Ler com cabeçalho correto
    df = pd.read_excel(temp_input.name, header=linha_cabecalho)

    # Remover linhas vazias
    df = df.dropna(subset=["ALUNO"])

    # Remover colunas desnecessárias
    df = df.loc[:, ~df.columns.str.contains("Unnamed")]
    df = df.drop(columns=["SITUAÇÃO", "TOTAL"], errors="ignore")

    # Manter somente notas numéricas
    colunas_final = ["ALUNO"]

    for col in df.columns:
        if col == "ALUNO":
            continue

        # Testa se é coluna numérica:
        serie = pd.to_numeric(df[col], errors="coerce")

        # Se pelo menos metade da coluna for número, mantemos
        if serie.notna().sum() >= len(serie) * 0.8:
            colunas_final.append(col)

    df_final = df[colunas_final]

    st.subheader("📄 Resultado Final (somente notas):")
    st.dataframe(df_final)

    # Salvar arquivo final
    temp_out = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    df_final.to_excel(temp_out.name, index=False)

    with open(temp_out.name, "rb") as f:
        st.download_button(
            "⬇️ Baixar Planilha Somente com Notas",
            data=f.read(),
            file_name="notas_limpas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
