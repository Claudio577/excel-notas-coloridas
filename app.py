import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

# ---------------------------------------------------------
# FUNÇÃO ROBUSTA PARA LIMPAR PLANILHA DO MAPÃO
# ---------------------------------------------------------
def limpar_planilha(df):

    # --- 1) Normalização total (transforma tudo em string simples) ---
    def normalizar(x):
        if isinstance(x, (list, tuple, set, dict, np.ndarray, pd.Series, pd.DataFrame)):
            return ""
        if pd.isna(x):
            return ""
        return str(x).strip()

    df = df.applymap(normalizar)

    # Remove colunas e linhas vazias
    df = df.loc[:, (df != "").any(axis=0)]
    df = df[(df != "").any(axis=1)]

    # --- 2) Encontrar linha que contém "ALUNO" ---
    linha_cab = None
    for i, row in df.iterrows():
        if any("aluno" == str(c).lower() for c in row):
            linha_cab = i
            break

    if linha_cab is None:
        st.error("❌ Não encontrei a linha com 'ALUNO' na planilha.")
        return None

    df.columns = df.iloc[linha_cab].tolist()
    df = df.iloc[linha_cab + 1:]

    # Remove colunas Unnamed
    df = df.loc[:, ~df.columns.astype(str).str.contains("unnamed", case=False)]

    # --- 3) Converter automaticamente colunas que contêm números ---
    colunas_finais = ["ALUNO"]

    for col in df.columns:
        if col == "ALUNO":
            continue

        valores_convertidos = []
        for v in df[col]:
            try:
                valores_convertidos.append(float(v))
            except:
                valores_convertidos.append(np.nan)

        serie = pd.Series(valores_convertidos)

        # coluna é nota se houver pelo menos 1 número válido
        if serie.notna().sum() > 0:
            df[col] = serie
            colunas_finais.append(col)

    df = df[colunas_finais]

    return df


# ---------------------------------------------------------
# FUNÇÃO PARA JUNTAR B1, B2, B3 (ALUNO = chave)
# ---------------------------------------------------------
def juntar_bimestres(dfs):
    df_final = dfs[0]
    for i in range(1, len(dfs)):
        df_final = df_final.merge(dfs[i], on="ALUNO", how="outer")
    return df_final


# ---------------------------------------------------------
# FUNÇÃO PARA GERAR ARQUIVO EXCEL COM NOTAS VERMELHAS
# ---------------------------------------------------------
def gerar_excel_colorido(df):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Notas")

        workbook = writer.book
        worksheet = writer.sheets["Notas"]

        # formato vermelho
        red_format = workbook.add_format({'font_color': 'red'})

        # aplicar coloração
        for row in range(1, len(df) + 1):
            for col in range(1, len(df.columns)):
                try:
                    valor = df.iloc[row - 1, col]
                    if not pd.isna(valor) and valor < 5:
                        worksheet.write(row, col, valor, red_format)
                except:
                    pass

    output.seek(0)
    return output


# ---------------------------------------------------------
# INTERFACE STREAMLIT
# ---------------------------------------------------------
st.title("📘 Unificação de Notas – 1º, 2º e 3º Bimestre")
st.write("Envie as 3 planilhas do MAPÃO (1°, 2° e 3° bimestre).")

uploaded_b1 = st.file_uploader("1º Bimestre", type=["xlsx"])
uploaded_b2 = st.file_uploader("2º Bimestre", type=["xlsx"])
uploaded_b3 = st.file_uploader("3º Bimestre", type=["xlsx"])

if uploaded_b1 and uploaded_b2 and uploaded_b3:

    st.success("✔ Arquivos carregados! Processando...")

    try:
        # ler sem header
        df1 = limpar_planilha(pd.read_excel(uploaded_b1, header=None))
        df2 = limpar_planilha(pd.read_excel(uploaded_b2, header=None))
        df3 = limpar_planilha(pd.read_excel(uploaded_b3, header=None))

        if df1 is None or df2 is None or df3 is None:
            st.error("Erro ao ler alguma planilha.")
            st.stop()

        # renomear colunas para diferenciar bimestres
        df1 = df1.rename(columns={c: f"{c}_B1" for c in df1.columns if c != "ALUNO"})
        df2 = df2.rename(columns={c: f"{c}_B2" for c in df2.columns if c != "ALUNO"})
        df3 = df3.rename(columns={c: f"{c}_B3" for c in df3.columns if c != "ALUNO"})

        # unir
        df_final = juntar_bimestres([df1, df2, df3])

        st.subheader("📄 Planilha Final (antes da coloração)")
        st.dataframe(df_final)

        # gerar excel colorido
        excel_file = gerar_excel_colorido(df_final)

        st.download_button(
            "⬇ Baixar Planilha Final (Notas < 5 em Vermelho)",
            data=excel_file,
            file_name="notas_unificadas_coloridas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error("❌ Erro inesperado ao processar.")
        st.exception(e)
