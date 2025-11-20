import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill

st.title("📘 Unificador de Notas - MAPÃO (3 Bimestres)")


# ----------------------------------------------------------
# FUNÇÃO DE LIMPEZA DEFINITIVA PARA SEU MAPÃO
# ----------------------------------------------------------
def limpar_planilha(df):

    # 🔥 Função para achatar qualquer objeto estranho na célula
    def normalizar_celula(x):
        if isinstance(x, (list, tuple)):
            return " ".join(map(str, x))
        if isinstance(x, pd.DataFrame):
            return str(x.values.flatten()[0]) if len(x.values.flatten()) else ""
        if isinstance(x, pd.Series):
            return str(x.iloc[0]) if len(x) else ""
        return str(x) if pd.notna(x) else ""

    # Aplica normalização
    df = df.applymap(normalizar_celula)

    # Remove células com "nan"
    df = df.replace("nan", "")
    df = df.replace("None", "")

    # Remove colunas e linhas totalmente vazias
    df = df.loc[:, (df != "").any(axis=0)]
    df = df[(df != "").any(axis=1)]

    # Encontra a linha do cabeçalho real (ALUNO)
    linha_aluno = None
    for i, row in df.iterrows():
        if row.astype(str).str.contains("ALUNO", case=False).any():
            linha_aluno = i
            break

    if linha_aluno is None:
        st.error("❌ Não encontrei a linha com 'ALUNO'.")
        return None

    df.columns = df.iloc[linha_aluno]
    df = df.iloc[linha_aluno + 1:]

    # Remove Unnamed
    df = df.loc[:, ~df.columns.astype(str).str.contains("Unnamed")]

    # Identificar colunas com números
    colunas_ok = ["ALUNO"]
    for col in df.columns:
        if col == "ALUNO":
            continue

        test = pd.to_numeric(df[col], errors="coerce")

        # Se houver pelo menos 1 número, aceita
        if test.notna().sum() > 0:
            colunas_ok.append(col)

    df = df[colunas_ok]

    # Converter notas para número
    for col in df.columns:
        if col != "ALUNO":
            df[col] = pd.to_numeric(df[col], errors="coerce")

    return df


# ----------------------------------------------------------
# UPLOADS
# ----------------------------------------------------------
uploaded_b1 = st.file_uploader("📥 Envie o 1º Bimestre", type=["xlsx"])
uploaded_b2 = st.file_uploader("📥 Envie o 2º Bimestre", type=["xlsx"])
uploaded_b3 = st.file_uploader("📥 Envie o 3º Bimestre", type=["xlsx"])

if uploaded_b1 and uploaded_b2 and uploaded_b3:

    st.success("✔ Arquivos carregados! Processando...")

    df1 = limpar_planilha(pd.read_excel(uploaded_b1, header=None))
    df2 = limpar_planilha(pd.read_excel(uploaded_b2, header=None))
    df3 = limpar_planilha(pd.read_excel(uploaded_b3, header=None))

    if df1 is None or df2 is None or df3 is None:
        st.stop()

    # Renomear colunas por bimestre
    df1 = df1.rename(columns={c: f"{c}_B1" for c in df1.columns if c != "ALUNO"})
    df2 = df2.rename(columns={c: f"{c}_B2" for c in df2.columns if c != "ALUNO"})
    df3 = df3.rename(columns={c: f"{c}_B3" for c in df3.columns if c != "ALUNO"})

    # Unificar
    df_final = df1.merge(df2, on="ALUNO", how="left").merge(df3, on="ALUNO", how="left")

    st.subheader("📄 Prévia")
    st.dataframe(df_final)

    # ----------------------------------------------------------
    # GERAR ARQUIVO COLORIDO
    # ----------------------------------------------------------
    def exportar(df):
        wb = Workbook()
        ws = wb.active
        ws.append(df.columns.tolist())

        vermelho = PatternFill(start_color="FF6666", end_color="FF6666", fill_type="solid")

        for _, row in df.iterrows():
            ws.append(row.tolist())

        for row in ws.iter_rows(min_row=2):
            for cell in row:
                if isinstance(cell.value, (int, float)) and cell.value < 5:
                    cell.fill = vermelho

        return wb

    wb = exportar(df_final)

    def baixar():
        from io import BytesIO
        buffer = BytesIO()
        wb.save(buffer)
        buffer.seek(0)
        return buffer

    st.download_button(
        "⬇ Baixar Notas Unificadas (Vermelhas <5)",
        data=baixar(),
        file_name="notas_unificadas.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
