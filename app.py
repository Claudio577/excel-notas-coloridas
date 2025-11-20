import streamlit as st
import pandas as pd
import numpy as np
import tempfile
import re
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter


st.title("📘 Unificador de Notas – 1º, 2º e 3º Bimestres (Notas Vermelhas < 5)")


# --------------------------------------------------------------
#  FUNÇÃO PARA DETECTAR SE O TEXTO É UM NOME DE ALUNO REAL
# --------------------------------------------------------------

def eh_aluno(nome):
    if pd.isna(nome):
        return False

    partes = str(nome).split()

    if len(partes) < 2:
        return False

    if not all(p.isalpha() for p in partes):
        return False

    if len(partes[0]) <= 2:  # remove EP, ES, ET, AC...
        return False

    return True


# --------------------------------------------------------------
#  LIMPEZA DAS PLANILHAS
# --------------------------------------------------------------

def limpar_planilha(file):
    df_raw = pd.read_excel(file, header=None)

    try:
        linha_cab = df_raw[df_raw.iloc[:, 0] == "ALUNO"].index[0]
    except:
        st.error("❌ A planilha enviada não contém a coluna 'ALUNO'.")
        st.stop()

    df = pd.read_excel(file, header=linha_cab)

    df = df[df["ALUNO"].apply(eh_aluno)]

    df = df.loc[:, ~df.columns.str.contains("Unnamed")]
    df = df.drop(columns=["SITUAÇÃO", "TOTAL"], errors="ignore")

    materias_proibidas = ["arte", "esporte", "música", "artes", "big a", "inovação", "inovacao"]

    def coluna_proibida(col):
        texto = col.lower()
        return any(p in texto for p in materias_proibidas)

    df = df[[c for c in df.columns if not coluna_proibida(c)]]

    def extrair_nota(valor):
        if pd.isna(valor):
            return np.nan
        # A regex original não lida bem com notas como "10,0" ou "9.5" se o input for string,
        # mas como o exemplo parece ter notas inteiras, vou manter o que está perto do original,
        # focando apenas no primeiro número inteiro de 0 a 10.
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

        if df[col].notna().sum() > 0:
            colunas_validas.append(col)
        else:
            continue

        materia = re.split(r"\d+", col)[0].strip().lower()

        renomear[col] = materia

    df = df[colunas_validas]
    df = df.rename(columns=renomear)

    return df


# --------------------------------------------------------------
#  FORMATAÇÃO DO CABEÇALHO EM 2 LINHAS
#  A CORREÇÃO DE "SEGUNDA LINHA INVERTIDA" E A REMOÇÃO DE LINHAS ESTÃO AQUI.
# --------------------------------------------------------------

def formatar_cabecalho_simples(path, df_final):
    wb = load_workbook(path)
    ws = wb.active

    # 1. REMOVER LINHAS 3, 4 E 5 (Antes de inserir o novo cabeçalho)
    # Como o DataFrame foi escrito com startrow=3, as linhas 3, 4 e 5 da planilha
    # são: A linha em branco do startrow=3, a linha do cabeçalho do DF (A4)
    # e a primeira linha de dados (A5). Queremos manter apenas os dados.
    # Os dados começam na linha 4 (depois do startrow=3).
    # Vamos deletar as 3 primeiras linhas: 1, 2 e 3.
    ws.delete_rows(1, 3) 
    
    # Após deletar as 3 primeiras linhas, a primeira linha de dados
    # agora está na Linha 1 do Excel. Vamos inserir as 2 linhas
    # para o novo cabeçalho (MATÉRIA e BIMESTRE).
    ws.insert_rows(1)
    ws.insert_rows(2)

    # 2. ESCREVER NOVO CABEÇALHO
    ws["A1"] = "ALUNO"
    ws["A2"] = ""

    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

    col_excel = 2
    
    # 3. CORREÇÃO DA ORDEM DO BIMESTRE INVERTIDO
    colunas_agrupadas = {}
    
    # O DataFrame final já está na ordem correta, mas precisamos garantir
    # que a iteração aqui siga essa ordem para escrever corretamente.
    # O cabeçalho no Excel deve seguir a ordem das colunas do DF.
    
    for col in df_final.columns:
        if col == "ALUNO":
            continue

        # A coluna no df_final tem o formato "materia_BI"
        partes = col.split("_")
        materia = partes[0]
        bi = partes[1]

        # Escreve a Matéria na Linha 1
        ws.cell(row=1, column=col_excel, value=materia.capitalize())

        # Escreve o Bimestre na Linha 2
        bimestre_formatado = bi.replace("B", "ºBi")  # Ex: B1 → 1ºBi
        ws.cell(row=2, column=col_excel, value=bimestre_formatado)
        
        col_excel += 1

    wb.save(path)


# --------------------------------------------------------------
#  UPLOAD DOS 3 BIMESTRES
# --------------------------------------------------------------

file_b1 = st.file_uploader("📤 Envie o Excel do 1º Bimestre", type=["xlsx"])
file_b2 = st.file_uploader("📤 Envie o Excel do 2º Bimestre", type=["xlsx"])
file_b3 = st.file_uploader("📤 Envie o Excel do 3º Bimestre", type=["xlsx"])

if file_b1 and file_b2 and file_b3:

    st.success("Arquivos carregados! Processando...")

    df1 = limpar_planilha(file_b1)
    df2 = limpar_planilha(file_b2)
    df3 = limpar_planilha(file_b3)

    ordem_b1 = df1["ALUNO"].tolist()

    df1 = df1.rename(columns={c: f"{c}_B1" for c in df1.columns if c != "ALUNO"})
    df2 = df2.rename(columns={c: f"{c}_B2" for c in df2.columns if c != "ALUNO"})
    df3 = df3.rename(columns={c: f"{c}_B3" for c in df3.columns if c != "ALUNO"})

    df_final = df1.merge(df2, on="ALUNO", how="outer")
    df_final = df_final.merge(df3, on="ALUNO", how="outer")

    df_final = df_final.fillna("–")

    df_final["ordem"] = df_final["ALUNO"].apply(
        lambda nome: ordem_b1.index(nome) if nome in ordem_b1 else 999
    )
    df_final = df_final.sort_values("ordem").drop(columns=["ordem"])

    st.subheader("📄 Planilha Final (antes da coloração)")
    st.dataframe(df_final)

    temp_out = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    # O startrow=3 cria uma linha em branco (L1), o cabeçalho do DF (L2) e os dados (L3...)
    # Usaremos a função de formatação para remover as linhas iniciais indesejadas (L1, L2, L3)
    # e depois inserir o cabeçalho correto nas novas L1 e L2.
    df_final.to_excel(temp_out.name, index=False, startrow=0) # startrow=0 para começar na L1

    formatar_cabecalho_simples(temp_out.name, df_final)

    def colorir_notas(path):
        wb = load_workbook(path)
        ws = wb.active
        red = Font(color="FF0000", bold=True)

        # Começa a colorir a partir da Linha 3 (onde os dados do aluno começam agora)
        for col in range(2, ws.max_column + 1):
            for row in range(3, ws.max_row + 1):
                val = ws.cell(row=row, column=col).value
                try:
                    if isinstance(val, (int, float)) and val < 5:
                        ws.cell(row=row, column=col).font = red
                except:
                    pass

        wb.save(path)

    colorir_notas(temp_out.name)

    with open(temp_out.name, "rb") as f:
        st.download_button(
            "⬇️ Baixar Planilha Final (Formatada + Notas Vermelhas)",
            f.read(),
            file_name="notas_unificadas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

