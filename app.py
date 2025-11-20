import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import tempfile

st.title("📊 Colorizador de Notas em Excel (Notas Baixas em Vermelho)")

st.write("""
Envie um arquivo Excel (.xlsx) onde:
- A primeira coluna é **aluno**
- As demais colunas são **notas das matérias**
- Notas abaixo de 6 serão destacadas em **vermelho**
""")

uploaded_file = st.file_uploader("Envie seu Excel:", type=["xlsx"])

if uploaded_file:

    # Salvar arquivo temporário no servidor do Streamlit
    temp_input = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    temp_input.write(uploaded_file.getbuffer())
    temp_input.close()

    # Ler o Excel
    df = pd.read_excel(temp_input.name)

    st.subheader("📄 Prévia do arquivo enviado:")
    st.dataframe(df)

    # Definir limite da nota
    nota_limite = st.number_input(
        "Considere nota baixa abaixo de:",
        min_value=0.0, max_value=10.0, value=5.0
    )

    if st.button("🎨 Colorir notas baixas"):
        # Salvar saída temporária
        temp_output = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")

        # Salvar o df para edição posterior
        df.to_excel(temp_output.name, index=False)

        # Abrir no openpyxl para pintar
        wb = load_workbook(temp_output.name)
        ws = wb.active

        vermelho = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")

        # Iterar sobre as células a partir da 2ª linha e 2ª coluna
        for row in ws.iter_rows(min_row=2, min_col=2):
            for cell in row:
                try:
                    valor = float(cell.value)
                    if valor < nota_limite:
                        cell.fill = vermelho
                except:
                    pass  # ignora valores não numéricos

        wb.save(temp_output.name)

        st.success("Arquivo gerado com sucesso! Baixe abaixo:")

        with open(temp_output.name, "rb") as f:
            st.download_button(
                "⬇️ Baixar Excel Colorido",
                data=f,
                file_name="notas_coloridas.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
