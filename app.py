import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill

st.title("📘 Unificador de Notas - 3 Bimestres (Com Notas <5 em Vermelho)")

# ----------------------------
# FUNÇÃO DE LIMPAR PLANILHAS
# ----------------------------

def limpar_planilha(df):
    # Remove linhas inúteis
    df = df.dropna(how="all")
    df = df.dropna(axis=1, how="all")

    # Acha a linha onde começa a lista de alunos
    linha_alunos = df[df.astype(str).apply(lambda row: row.str.contains("ALUNO", case=False)).any(axis=1)].index[0]
    df.columns = df.iloc[linha_alunos]
    df = df.iloc[linha_alunos + 1:]

    # Remove colunas que não são notas
    df = df[df.columns.dropna()]
    df = df.loc[:, ~df.columns.astype(str).str.contains("Unnamed")]

    # Mantém apenas colunas com ALUNO e números
    cols_ok = ["ALUNO"] + [c for c in df.columns if df[c].astype(str).str.replace('.', '').str.isdigit().any()]
    df = df[cols_ok]

    # Converte notas para número
    for col in df.columns:
        if col != "ALUNO":
            df[col] = pd.to_numeric(df[col], errors="coerce")

    return df

# ----------------------------
# PROCESSAMENTO PRINCIPAL
# ----------------------------

uploaded_b1 = st.file_uploader("📥 Envie o arquivo do 1º Bimestre", type=["xlsx"])
uploaded_b2 = st.file_uploader("📥 Envie o arquivo do 2º Bimestre", type=["xlsx"])
uploaded_b3 = st.file_uploader("📥 Envie o arquivo do 3º Bimestre", type=["xlsx"])

if uploaded_b1 and uploaded_b2 and uploaded_b3:
    st.success("✔ Arquivos carregados!")

    df1 = limpar_planilha(pd.read_excel(uploaded_b1))
    df2 = limpar_planilha(pd.read_excel(uploaded_b2))
    df3 = limpar_planilha(pd.read_excel(uploaded_b3))

    # Renomeia colunas adicionando o bimestre
    df1 = df1.rename(columns={c: f"{c}_B1" for c in df1.columns if c != "ALUNO"})
    df2 = df2.rename(columns={c: f"{c}_B2" for c in df2.columns if c != "ALUNO"})
    df3 = df3.rename(columns={c: f"{c}_B3" for c in df3.columns if c != "ALUNO"})

    # Junta tudo
    df_final = df1.merge(df2, on="ALUNO", how="left").merge(df3, on="ALUNO", how="left")

    st.subheader("📄 Prévia (antes da coloração)")
    st.dataframe(df_final)

    # ----------------------------
    # EXPORTAR COM AS NOTAS < 5 EM VERMELHO
    # ----------------------------

    def exportar_excel_colorido(df):
        wb = Workbook()
        ws = wb.active

        # Escreve cabeçalhos
        ws.append(df.columns.tolist())

        red_fill = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")

        # Escreve dados com coloração
        for row in df.itertuples(index=False):
            linha = []
            for i, valor in enumerate(row):
                linha.append(valor)
            ws.append(linha)

        # Aplica cor nas notas < 5
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                try:
                    if isinstance(cell.value, (int, float)) and cell.value < 5:
                        cell.fill = red_fill
                except:
                    pass

        return wb

    wb = exportar_excel_colorido(df_final)

    st.download_button(
        "⬇ Baixar Planilha Final (Notas <5 em Vermelho)",
        data=lambda: wb.save("notas_unificadas.xlsx"),
        file_name="notas_unificadas.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

