import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill

st.title("📘 Unificador de Notas - 3 Bimestres (Notas <5 em Vermelho)")


# ----------------------------------------------------------
# FUNÇÃO: LIMPA PLANILHA DE ACORDO COM SEU MODELO REAL
# ----------------------------------------------------------
def limpar_planilha(df):

    # Remove linhas e colunas completamente vazias
    df = df.dropna(how="all")
    df = df.dropna(axis=1, how="all")

    # Encontrar a linha onde está "ALUNO"
    linha_aluno = None
    for i, row in df.iterrows():
        if row.astype(str).str.contains("ALUNO", case=False).any():
            linha_aluno = i
            break

    if linha_aluno is None:
        st.error("Não encontrei a linha com 'ALUNO'. Verifique a planilha.")
        return None

    # Define cabeçalhos reais
    df.columns = df.iloc[linha_aluno]
    df = df.iloc[linha_aluno + 1:]

    # Remove colunas inválidas
    df = df.loc[:, ~df.columns.astype(str).str.contains("Unnamed")]

    # Agora identificar SOMENTE colunas com números
    colunas_numericas = ["ALUNO"]

    for col in df.columns:
        if col == "ALUNO":
            continue

        # Tenta converter valores para número
        valores_convertidos = pd.to_numeric(df[col], errors="coerce")

        # Se pelo menos 1 valor virar número → é uma nota
        if valores_convertidos.notna().sum() > 0:
            colunas_numericas.append(col)

    # mantém só ALUNO + notas
    df = df[colunas_numericas]

    # Converte notas para número de verdade
    for col in df.columns:
        if col != "ALUNO":
            df[col] = pd.to_numeric(df[col], errors="coerce")

    return df


# ----------------------------------------------------------
# INTERFACE STREAMLIT
# ----------------------------------------------------------

uploaded_b1 = st.file_uploader("📥 Envie o 1º Bimestre", type=["xlsx"])
uploaded_b2 = st.file_uploader("📥 Envie o 2º Bimestre", type=["xlsx"])
uploaded_b3 = st.file_uploader("📥 Envie o 3º Bimestre", type=["xlsx"])

if uploaded_b1 and uploaded_b2 and uploaded_b3:

    st.success("✔ Arquivos carregados! Processando...")

    df1 = limpar_planilha(pd.read_excel(uploaded_b1))
    df2 = limpar_planilha(pd.read_excel(uploaded_b2))
    df3 = limpar_planilha(pd.read_excel(uploaded_b3))

    if df1 is None or df2 is None or df3 is None:
        st.stop()

    # Renomeia colunas por bimestre
    df1 = df1.rename(columns={c: f"{c}_B1" for c in df1.columns if c != "ALUNO"})
    df2 = df2.rename(columns={c: f"{c}_B2" for c in df2.columns if c != "ALUNO"})
    df3 = df3.rename(columns={c: f"{c}_B3" for c in df3.columns if c != "ALUNO"})

    # Junta tudo
    df_final = df1.merge(df2, on="ALUNO", how="left").merge(df3, on="ALUNO", how="left")

    st.subheader("📄 Prévia antes da coloração")
    st.dataframe(df_final)

    # ----------------------------------------------------------
    # EXPORTAR COM COLORAÇÃO
    # ----------------------------------------------------------

    def exportar_colorido(df):
        wb = Workbook()
        ws = wb.active

        ws.append(df.columns.tolist())

        vermelho = PatternFill(start_color="FF6666", end_color="FF6666", fill_type="solid")

        for index, row in df.iterrows():
            ws.append(row.tolist())

        # pinta notas
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                try:
                    if isinstance(cell.value, (int, float)) and cell.value < 5:
                        cell.fill = vermelho
                except:
                    pass

        return wb

    wb = exportar_colorido(df_final)

    def salvar_wb():
        from io import BytesIO
        buffer = BytesIO()
        wb.save(buffer)
        buffer.seek(0)
        return buffer

    st.download_button(
        "⬇ Baixar Planilha Final (Notas <5 em Vermelho)",
        data=salvar_wb(),
        file_name="notas_unificadas.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

