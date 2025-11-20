import streamlit as st
import pandas as pd
import numpy as np
import tempfile
import re
from openpyxl import load_workbook
from openpyxl.styles import Font

st.title("📘 Unificador de Notas – 1º, 2º e 3º Bimestres (Notas Vermelhas < 5)")

st.write("""
Envie os 3 arquivos, um para cada bimestre.  
O sistema irá:
- Limpar os dados  
- Remover colunas inúteis  
- Extrair somente notas (0–10)  
- Renomear matérias (remover códigos)  
- Unir tudo em uma única planilha  
- Pintar notas menores que 5 em vermelho  
""")

# ---------------- FUNÇÃO DE LIMPEZA ----------------

def limpar_planilha(file):
    df_raw = pd.read_excel(file, header=None)

    # Encontrar linha onde está escrito "ALUNO"
    linha_cabecalho = df_raw[df_raw.iloc[:, 0] == "ALUNO"].index[0]

    df = pd.read_excel(file, header=linha_cabecalho)

    df = df.dropna(subset=["ALUNO"])
    df = df.loc[:, ~df.columns.str.contains("Unnamed")]
    df = df.drop(columns=["SITUAÇÃO", "TOTAL"], errors="ignore")

    # Extrair número válido (0-10)
    def extrair_nota(valor):
        if pd.isna(valor):
            return np.nan
        nums = re.findall(r"\d+", str(valor))
        if not nums:
            return np.nan
        num = int(nums[0])
        if 0 <= num <= 10:
            return num
        return np.nan

    colunas_remover = []

    for col in df.columns:
        if col == "ALUNO":
            continue

        df[col] = df[col].apply(extrair_nota)

        # Remove coluna inteira se não tiver nenhuma nota
        if df[col].dropna().empty:
            colunas_remover.append(col)

    df = df.drop(columns=colunas_remover, errors="ignore")

    # Renomear colunas (remover códigos)
    def limpar_nome_coluna(nome):
        base = re.split(r"\d+", nome)[0].strip()
        return base if base else nome

    df.columns = [limpar_nome_coluna(col) for col in df.columns]

    return df


# ---------------- UPLOAD DOS 3 BIMESTRES ----------------

file_b1 = st.file_uploader("📤 Envie o Excel do 1º Bimestre", type=["xlsx"])
file_b2 = st.file_uploader("📤 Envie o Excel do 2º Bimestre", type=["xlsx"])
file_b3 = st.file_uploader("📤 Envie o Excel do 3º Bimestre", type=["xlsx"])

if file_b1 and file_b2 and file_b3:

    st.success("Arquivos carregados! Processando...")

    df1 = limpar_planilha(file_b1)
    df2 = limpar_planilha(file_b2)
    df3 = limpar_planilha(file_b3)

    # Renomear colunas
    df1 = df1.rename(columns={col: f"{col}_B1" for col in df1.columns if col != "ALUNO"})
    df2 = df2.rename(columns={col: f"{col}_B2" for col in df2.columns if col != "ALUNO"})
    df3 = df3.rename(columns={col: f"{col}_B3" for col in df3.columns if col != "ALUNO"})

    # Unir tudo por ALUNO
    df_final = df1.merge(df2, on="ALUNO", how="outer")
    df_final = df_final.merge(df3, on="ALUNO", how="outer")

    st.subheader("📄 Planilha Final (antes da coloração)")
    st.dataframe(df_final)

    # ---------------- GERAR EXCEL COM NOTAS < 5 EM VERMELHO ----------------

    temp_out = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    df_final.to_excel(temp_out.name, index=False)

    def colorir_notas(path):
        wb = load_workbook(path)
        ws = wb.active
        red_font = Font(color="FF0000", bold=True)

        # Começa na coluna 2 porque a 1 é "ALUNO"
        for col in range(2, ws.max_column + 1):
            for row in range(2, ws.max_row + 1):
                cell = ws.cell(row=row, column=col)
                try:
                    if isinstance(cell.value, (int, float)) and cell.value < 5:
                        cell.font = red_font
                except:
                    pass

        wb.save(path)

    colorir_notas(temp_out.name)

    with open(temp_out.name, "rb") as f:
        st.download_button(
            "⬇️ Baixar Planilha Final Unificada (Notas <5 em Vermelho)",
            f.read(),
            file_name="notas_bimestres_unificadas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
