import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import tempfile

st.title("📊 Colorizador de Notas (Corrigido: ignora M e F)")

st.write("""
Este aplicativo:
- Ignora linhas de 'M' e 'F' das matérias
- Só colore NOTAS dos alunos (linha 14 em diante)
- Pinta de vermelho notas abaixo do limite
""")

uploaded_file = st.file_uploader("Envie seu Excel:", type=["xlsx"])

if uploaded_file:

    # Salvar arquivo temporário
    temp_input = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    temp_input.write(uploaded_file.getbuffer())
    temp_input.close()

    # Ler a planilha como DataFrame apenas para pré-visualização
    df = pd.read_excel(temp_input.name, header=None)
    
    st.subheader("📄 Prévia do arquivo:")
    st.dataframe(df)

    nota_limite = st.number_input(
        "Considere nota baixa abaixo de:",
        min_value=0.0, max_value=10.0, value=6.0
    )

    if st.button("🎨 Colorir Notas"):
        temp_output = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")

        # Abrir no openpyxl
        wb = load_workbook(temp_input.name)
        ws = wb.active

        vermelho = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")

        # 🔥 IMPORTANTE:
        # Notas começam NA LINHA 14 (1-based)
        primeira_linha_nota = 14

        # Iterar a partir da linha 14
        for row in ws.iter_rows(min_row=primeira_linha_nota, min_col=2):
            for cell in row:
                try:
                    valor = float(cell.value)
                    if valor < nota_limite:
                        cell.fill = vermelho
                except:
                    # Se não for número, ignora (é M, F, AC, ES etc)
                    pass

        wb.save(temp_output.name)

        st.success("Arquivo gerado! Baixe abaixo:")

        with open(temp_output.name, "rb") as f:
            st.download_button(
                "⬇️ Baixar Excel Colorido",
                data=f,
                file_name="notas_coloridas.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
