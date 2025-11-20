import streamlit as st
import pandas as pd
import tempfile

st.title("📘 Extrator de Notas M (Planilha Escolar)")

st.write("""
Este aplicativo extrai automaticamente:
- Os **nomes dos alunos**
- As **notas M** de cada matéria (linha onde aparece o 'M')
- Ignora F, AC, ES e todas as outras linhas
""")

uploaded_file = st.file_uploader("Envie o arquivo Excel (.xlsx):", type=["xlsx"])

if uploaded_file:

    # Salvar temporariamente
    temp_input = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    temp_input.write(uploaded_file.getbuffer())
    temp_input.close()

    # Ler o Excel SEM cabeçalho
    df = pd.read_excel(temp_input.name, header=None)

    st.subheader("Prévia das primeiras linhas:")
    st.dataframe(df.head(20))

    st.info("""
    Identificação automática:
    - Linha 11 → Nomes das matérias  
    - Linha 12 → Notas M  
    - Linha 14 em diante → Nomes dos alunos  
    """)

    # Linha 11 → cabeçalho das matérias (index 10)
    materias = df.iloc[10].values

    # Linha 12 → notas M (index 11)
    notas_M = df.iloc[11].values

    # Linhas 14 → nomes dos alunos (index 13)
    alunos = df.iloc[13:, 0].values

    # Criar DataFrame final
    notas_expandidas = pd.DataFrame({"aluno": alunos})

    # Montar coluna por coluna
    for i in range(1, len(materias)):  # coluna 1 em diante
        materia = str(materias[i]).strip()

        # Corrigir se vier NaN
        if materia == "nan":
            continue

        nota = notas_M[i]

        # Converter nota para número
        try:
            nota = float(nota)
        except:
            nota = ""

        notas_expandidas[materia] = nota

    st.subheader("📄 Resultado Final – Notas M:")
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
