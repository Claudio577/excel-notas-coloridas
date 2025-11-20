import streamlit as st
import pandas as pd
import tempfile

st.title("📘 Extrator de Notas – Versão Simplificada")

st.write("""
Este app extrai automaticamente:
- Os **nomes dos alunos**
- As **notas de todas as matérias**
- Remove colunas como **SITUAÇÃO**, **TOTAL**, etc.
""")

uploaded_file = st.file_uploader("Envie o Excel (.xlsx):", type=["xlsx"])

if uploaded_file:

    # Salvar temporariamente
    temp_input = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    temp_input.write(uploaded_file.getbuffer())
    temp_input.close()

    df = pd.read_excel(temp_input.name, header=None)

    st.subheader("Primeiras linhas do arquivo detectado:")
    st.dataframe(df.head(20))

    st.info("Processando alunos e notas...")

    # Linha onde começam os alunos = linha que tem o texto "ALUNO"
    linha_aluno = df[df.iloc[:, 0] == "ALUNO"].index[0]

    # A linha seguinte é o cabeçalho
    header_row = linha_aluno

    # O próximo bloco são os alunos
    dados = pd.read_excel(temp_input.name, header=header_row)

    # Remover linhas em branco
    dados = dados.dropna(subset=["ALUNO"])

    # Remover colunas que não queremos
    colunas_remover = ["SITUAÇÃO", "TOTAL", "None", "Unnamed: 1"]
    colunas_existentes = [c for c in colunas_remover if c in dados.columns]
    dados = dados.drop(columns=colunas_existentes, errors="ignore")

    st.subheader("📄 Resultado Final (alunos + notas):")
    st.dataframe(dados)

    # Download da planilha final
    temp_output = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    dados.to_excel(temp_output.name, index=False)

    with open(temp_output.name, "rb") as f:
        st.download_button(
            "⬇️ Baixar Planilha de Notas (limpa)",
            data=f.read(),
            file_name="notas_extraidas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
