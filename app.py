import streamlit as st
import pandas as pd
import numpy as np
import tempfile
import re
from openpyxl import load_workbook
from openpyxl.styles import Font

st.title("📘 Unificador de Notas – 1º, 2º e 3º Bimestres (Notas Vermelhas < 5)")


# ---------------------------------------------------------------------
#  FUNÇÃO PARA DETECTAR SE O TEXTO É UM NOME DE ALUNO REAL
# ---------------------------------------------------------------------

def eh_aluno(nome):
    if pd.isna(nome):
        return False

    partes = str(nome).split()

    # Tem pelo menos duas palavras
    if len(partes) < 2:
        return False

    # Cada parte contém letras
    if not all(p.isalpha() for p in partes):
        return False

    # Evitar abreviações tipo EP, ES, ET, AC
    if len(partes[0]) <= 2:
        return False

    return True


# ---------------------------------------------------------------------
#  LIMPEZA DA PLANILHA
# ---------------------------------------------------------------------

def limpar_planilha(file):
    df_raw = pd.read_excel(file, header=None)

    # Encontrar a linha onde aparece "ALUNO"
    try:
        linha_cab = df_raw[df_raw.iloc[:, 0] == "ALUNO"].index[0]
    except:
        st.error("❌ A planilha enviada não contém uma coluna 'ALUNO'.")
        st.stop()

    df = pd.read_excel(file, header=linha_cab)

    # Manter somente alunos válidos
    df = df[df["ALUNO"].apply(eh_aluno)]

    # Remover colunas 'Unnamed'
    df = df.loc[:, ~df.columns.str.contains("Unnamed")]

    # Remover colunas inúteis caso existam
    df = df.drop(columns=["SITUAÇÃO", "TOTAL"], errors="ignore")

    # Função para extrair apenas números (0–10)
    def extrair_nota(valor):
        if pd.isna(valor):
            return np.nan
        nums = re.findall(r"\d+", str(valor))
        if not nums:
            return np.nan
        n = int(nums[0])
        return n if 0 <= n <= 10 else np.nan

    colunas_validas = ["ALUNO"]
    renomear = {}

    for col in df.columns:
        if col == "ALUNO":
            continue

        df[col] = df[col].apply(extrair_nota)

        # Só manter coluna se houver pelo menos 1 nota numérica
        if df[col].notna().sum() > 0:
            colunas_validas.append(col)
        else:
            continue

        # Limpar o nome da coluna removendo números
        materia = re.split(r"\d+", col)[0].strip()
        if materia == "":
            materia = col

        renomear[col] = materia

    df = df[colunas_validas]
    df = df.rename(columns=renomear)

    return df


# ---------------------------------------------------------------------
#  UPLOAD DOS 3 BIMESTRES
# ---------------------------------------------------------------------

file_b1 = st.file_uploader("📤 Envie o Excel do 1º Bimestre", type=["xlsx"])
file_b2 = st.file_uploader("📤 Envie o Excel do 2º Bimestre", type=["xlsx"])
file_b3 = st.file_uploader("📤 Envie o Excel do 3º Bimestre", type=["xlsx"])

if file_b1 and file_b2 and file_b3:

    st.success("Arquivos carregados! Limpando dados...")

    # Limpar
    df1 = limpar_planilha(file_b1)
    df2 = limpar_planilha(file_b2)
    df3 = limpar_planilha(file_b3)

    # Guardar ordem original do 1º bimestre
    ordem_b1 = df1["ALUNO"].tolist()

    # Renomear colunas
    df1 = df1.rename(columns={c: f"{c}_B1" for c in df1.columns if c != "ALUNO"})
    df2 = df2.rename(columns={c: f"{c}_B2" for c in df2.columns if c != "ALUNO"})
    df3 = df3.rename(columns={c: f"{c}_B3" for c in df3.columns if c != "ALUNO"})

    # Unificar mantendo todos os alunos
    df_final = df1.merge(df2, on="ALUNO", how="outer")
    df_final = df_final.merge(df3, on="ALUNO", how="outer")

    # Preencher notas faltantes
    df_final = df_final.fillna("–")

    # Ordenar de acordo com o 1º bimestre (quem não existe vai para o final)
    df_final["ordem"] = df_final["ALUNO"].apply(
        lambda nome: ordem_b1.index(nome) if nome in ordem_b1 else 999
    )
    df_final = df_final.sort_values("ordem").drop(columns=["ordem"])

    # Mostrar antes da coloração
    st.subheader("📄 Planilha Final (antes da coloração)")
    st.dataframe(df_final)

    # ---------------------------------------------------------------------
    #  SALVAR ARQUIVO E COLORIR NOTAS < 5 EM VERMELHO
    # ---------------------------------------------------------------------

    temp_out = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    df_final.to_excel(temp_out.name, index=False)

    def colorir_notas(path):
        wb = load_workbook(path)
        ws = wb.active
        red = Font(color="FF0000", bold=True)

        for col in range(2, ws.max_column + 1):
            for row in range(2, ws.max_row + 1):
                val = ws.cell(row=row, column=col).value
                try:
                    if isinstance(val, (int, float)) and val < 5:
                        ws.cell(row=row, column=col).font = red
                except:
                    pass

        wb.save(path)

    colorir_notas(temp_out.name)

    # Botão de download
    with open(temp_out.name, "rb") as f:
        st.download_button(
            "⬇️ Baixar Planilha Final Unificada (Notas < 5 em Vermelho)",
            f.read(),
            file_name="notas_unificadas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

