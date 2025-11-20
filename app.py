def limpar_planilha(file):
    df_raw = pd.read_excel(file, header=None)

    # Encontrar linha onde está escrito "ALUNO"
    linha_cabecalho = df_raw[df_raw.iloc[:, 0] == "ALUNO"].index[0]

    df = pd.read_excel(file, header=linha_cabecalho)

    # Remover linhas que não são alunos (EP, ES, ET, AC etc.)
    df = df[df["ALUNO"].str.contains(" ", na=False)]  # só mantém quem tem nome completo

    # Remover colunas Unnamed
    df = df.loc[:, ~df.columns.str.contains("Unnamed")]

    # Remover SITUAÇÃO, TOTAL, etc.
    df = df.drop(columns=["SITUAÇÃO", "TOTAL"], errors="ignore")

    # Extrair números válidos 0–10
    def extrair_nota(valor):
        if pd.isna(valor):
            return np.nan
        nums = re.findall(r"\d+", str(valor))
        if not nums:
            return np.nan
        num = int(nums[0])
        return num if 0 <= num <= 10 else np.nan

    colunas_boas = ["ALUNO"]
    novas_colunas = {}

    for col in df.columns:
        if col == "ALUNO":
            continue

        df[col] = df[col].apply(extrair_nota)

        # Se a coluna tem pelo menos UMA nota, ela é válida
        if df[col].notna().sum() > 0:
            colunas_boas.append(col)
        else:
            continue

        # Limpar nome da matéria (remove números)
        materia = re.split(r"\d+", col)[0].strip()
        if not materia:
            materia = col

        novas_colunas[col] = materia

    df = df[colunas_boas]
    df = df.rename(columns=novas_colunas)

    return df

