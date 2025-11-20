import streamlit as st
import pandas as pd
import tempfile

st.title("📘 Extrator de Notas M (Planilha Escolar)")

st.write("""
Este aplicativo extrai:
- Os **nomes dos alunos**
- As **notas M** (as notas que ficam logo abaixo do nome da matéria)
- Ignora F, AC, ES e todas as linhas extras
""")

uploaded_file = st.file_uploader("Envie o arquivo Excel (.xlsx):", type=["xlsx"])

if uploaded_file:

    # Salvar temporariamente para leitura
    temp_input = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    temp_input.write(uploaded_file.getbuffer())
    temp_input.close()

    # Ler toda a planilha SEM cabeçalho (para pegar estrutura real)
    df = pd.read_excel(temp_input.name, header=None)

    st.subheader("Prévia da planilha enviada:")
    st.dataframe(df.head(20))

    st.info("""
    Extraindo dados...
    - Linha 11 → matérias  
    - Linha 12 → notas M  
    - Linha 14 em diante → nomes dos alunos
    """)

    # Linha 11 = cabeçalhos das matérias (index 10)
    materias = df.iloc[10].values

    # Linha 12 = notas M (index 11)
    notas_M = df.iloc[11].values

    # Linhas 14+ = alunos (index 13+)
    alunos = df.iloc[13:, 0].values  # nomes na primeira coluna

    # Criar tabela final:
    # Cada aluno recebe TODAS as notas M
    notas_expandidas = pd.DataFrame({
        "aluno": alunos
    })

    # Criar um dicionário: { matéria: notaM }
    dict_notas = {}
    for i in range(1, len(materias)):  # começa na coluna 1
        materia = materias[i]
        nota = notas_M[i]

        # Converter nota — se não for número, vira vazio
        try:
            nota = float(nota)
        except:
            nota = ""

        dict_notas[materia] = nota

    # Repetir as NOTAS M para cada aluno (porque são fixas)
    for materia, nota in dict_notas.items():
        notas_expandidas[materia] = nota

    st.subheader("📄 Resultado Final:")
    st.dataframe(notas_expandidas)

    # Download
    temp_output = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    notas_expandidas.to_excel(temp_output.name, index=False)

    with open(temp_output.name, "rb") as f:
        st.download_button(
            "⬇️ Baixar Planilha com Notas M",
            data=f.read(),
            file_name="notas_M_extraidas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
