import streamlit as st
import pandas as pd
import tempfile
from openpyxl import load_workbook
from openpyxl.styles import Font

st.title("📘 Planilha Final – Nomes + Notas (<5 em vermelho)")

uploaded_file = st.file_uploader("Envie o Excel final com nomes e notas:", type=["xlsx"])

if uploaded_file:

    # Ler o Excel diretamente
    df = pd.read_excel(uploaded_file)

    # Manter somente colunas com ALUNO ou números
    colunas_final = ["ALUNO"]

    for col in df.columns:
        if col == "ALUNO":
            continue
        # Se coluna tiver ao menos um número, consideramos nota
        if pd.to_numeric(df[col], errors="coerce").notna().sum() > 0:
            colunas_final.append(col)

    df_final = df[colunas_final]

    st.subheader("📄 Prévia da Planilha Final:")
    st.dataframe(df_final)

    # Salvar temporário
    temp_out = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    df_final.to_excel(temp_out.name, index=False)

    # ---- COLORIR NOTAS < 5 EM VERMELHO ----
    wb = load_workbook(temp_out.name)
    ws = wb.active
    red_font = Font(color="FF0000", bold=True)

    for row in range(2, ws.max_row + 1):
        for col in range(2, ws.max_column + 1):  # ignora ALUNO
            cell = ws.cell(row, col)
            try:
                if isinstance(cell.value, (int, float)) and cell.value < 5:
                    cell.font = red_font
            except:
                pass

    wb.save(temp_out.name)

    # Download
    with open(temp_out.name, "rb") as f:
        st.download_button(
            "⬇️ Baixar Planilha Final com Notas Vermelhas",
            f.read(),
            file_name="notas_formatadas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
