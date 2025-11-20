import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Unificador de Notas", layout="wide")

st.title("📘 Unificar Notas – 1º, 2º e 3º Bimestres")

st.write("Envie as três planilhas (1º, 2º e 3º bimestre).")


# ------------------------------------------------------------
# Função: encontrar linha onde começa ALUNO
# ------------------------------------------------------------
def encontrar_linha_aluno(df):
    for idx, row in df.iterrows():
        if row.astype(str).str.contains("ALUNO", case=False, regex=False).any():
            return idx
    return None


# ------------------------------------------------------------
# Limpar planilha individual
# ------------------------------------------------------------
def limpar_planilha(df_original, sufixo):
    linha = encontrar_linha_aluno(df_original)

    if linha is None:
        raise ValueError("Não foi encontrada a linha 'ALUNO' na planilha.")

    df = pd.read_excel(uploaded_file, header=linha)

    # Remover colunas vazias
    df = df.dropna(axis=1, how='all')

    # Remover linhas onde ALUNO está vazio ou é texto administrativo
    df = df[df["ALUNO"].astype(str).str.len() > 3]
    df = df[~df["ALUNO"].str.contains("Engajamento|Frequência|Compensada", case=False, na=False)]

    # Renomear colunas removendo números
    novas_colunas = {}
    for col in df.columns:
        novo = re.sub(r"\d+", "", col).strip().replace("  ", " ")
        novas_colunas[col] = f"{novo}_{sufixo}"

    novas_colunas["ALUNO"] = "ALUNO"  # manter nome original
    df = df.rename(columns=novas_colunas)

    # Converter notas para número
    for col in df.columns:
        if col != "ALUNO":
            df[col] = pd.to_numeric(df[col], errors="coerce")

    return df


# ------------------------------------------------------------
# Juntar bimestres
# ------------------------------------------------------------
def juntar_bimestres(df1, df2, df3):
    return df1.merge(df2, on="ALUNO", how="outer").merge(df3, on="ALUNO", how="outer")


# ------------------------------------------------------------
# Download com formatação de notas vermelhas
# ------------------------------------------------------------
def gerar_excel_colorido(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Notas', startrow=1)

    workbook = writer.book
    worksheet = writer.sheets['Notas']

    red_format = workbook.add_format({"font_color": "red"})

    # aplicar vermelho onde nota < 5
    for row in range(2, len(df) + 2):
        for col in range(1, len(df.columns)):
            val = df.iloc[row - 2, col]
            if pd.notna(val) and isinstance(val, (int, float)) and val < 5:
                worksheet.write(row, col, val, red_format)

    writer.save()
    return output.getvalue()


# ------------------------------------------------------------
# Uploads
# ------------------------------------------------------------
uploaded_b1 = st.file_uploader("📄 Envie o 1º Bimestre", type=["xlsx"])
uploaded_b2 = st.file_uploader("📄 Envie o 2º Bimestre", type=["xlsx"])
uploaded_b3 = st.file_uploader("📄 Envie o 3º Bimestre", type=["xlsx"])


if uploaded_b1 and uploaded_b2 and uploaded_b3:
    st.success("✔ Arquivos carregados! Processando...")

    # Processamento
    df1 = limpar_planilha(pd.read_excel(uploaded_b1, header=None), "B1")
    df2 = limpar_planilha(pd.read_excel(uploaded_b2, header=None), "B2")
    df3 = limpar_planilha(pd.read_excel(uploaded_b3, header=None), "B3")

    final = juntar_bimestres(df1, df2, df3)

    st.subheader("📘 Planilha Final (antes da coloração)")
    st.dataframe(final, height=500)

    # Arquivo para download
    excel_final = gerar_excel_colorido(final)

    st.download_button(
        label="⬇ Baixar Planilha Unificada (Notas <5 em Vermelho)",
        data=excel_final,
        file_name="notas_unificadas.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
