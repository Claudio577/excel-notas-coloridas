import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill

st.title("📘 Unificador de Notas - 3 Bimestres (Notas <5 em Vermelho)")


# ----------------------------------------------------------
# FUNÇÃO: LIMPA PLANILHA DE ACORDO COM SEU EXCEL REAL
# ----------------------------------------------------------
def limpar_planilha(df):

    # 🔥 Converte todo o DataFrame para texto simples
    df = df.applymap(lambda x: str(x) if not pd.isna(x) else "")

    # Remove linhas e colunas completamente vazias
    df = df.replace("nan", "")
    df = df.replace("None", "")

    df = df.loc[:, (df != "").any(axis=0)]
    df = df[(df != "").any(axis=1)]

    # Encontrar a linha onde está "ALUNO"
    linha_aluno = None
    for i, row in df.iterrows():
        if row.astype(str).str.contains("ALUNO", case=False).any():
            linha_aluno = i
            break

    if linha_aluno is None:
        st.error("❌ Não encontrei a linha com 'ALUNO'. Verifique a planilha.")
        return None

    # Cabeçalho real
    df.columns = df.iloc[linha_aluno]
    df = df.iloc[linha_aluno + 1:]

    # Remover colunas Unnamed
    df = df.loc[:, ~df.columns.astype(str).str.contains("Unnamed")]

    # Agora identificar SOMENTE colunas numéricas
    colunas_validas = ["ALUNO"]

    for col in df.columns:
        if col == "ALUNO":
            continue

        # Tenta converter para número
        numeros = pd.to_numeric(df[col], errors="coerce")

        # Se houver pelo menos 1 número, a coluna é nota
        if numeros.notna().sum() > 0:
            colunas_validas.append(col)

    # Mantém só as colunas de interesse
    df = df[colunas_validas]

    # Converte notas para número real
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

    df1 = limpar_planilha(pd.read_excel(uploaded_b1, header=None))
    df2 = limpar_planilha(pd.read_excel(uploaded_b2, header=None))
    df3 = limpar_planilha(pd.read_excel(uploaded_b3, header=None))

    if df1 is None or df2 is None or df3 is None:
        st.stop()

    # Renomeia colunas conforme o bimestre
    df1 = df1.rename(columns={c: f"{c}_B1" for c in df1.columns if c != "ALUNO"})
    df2 = df2.rename(columns={c: f"{c}_B2" for c in df2.columns if c != "ALUNO"})
    df3 = df3.rename(columns={c: f"{c}_B3" for c in df3.columns if c != "ALUNO"})

    # Junta os 3 bimestres
    df_final = df1.merge(df2, on="ALUNO", how="left").merge(df3, on="ALUNO", how="left")

    st.subheader("📄 Prévia antes da coloração")
    st.dataframe(df_final)


    # ----------------------------------------------------------
    # EXPORTAR ARQUIVO COLORIDO
    # ----------------------------------------------------------
    def exportar_colorido(df):
        wb = Workbook()
        ws = wb.active

        # Cabeçalho
        ws.append(df.columns.tolist())

        vermelho = PatternFill(start_color="FF6666", end_color="FF6666", fill_type="solid")

        # Preenche dados
        for _, row in df.iterrows():
            ws.append(row.tolist())

        # Pinta notas < 5
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
        "⬇ Baixar Planilha Final Unificada (Notas <5 Vermelho)",
        data=salvar_wb(),
        file_name="notas_unificadas.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

