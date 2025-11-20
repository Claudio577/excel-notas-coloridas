    # ---------------------------------------------------------------------
    #  SALVAR ARQUIVO E COLORIR NOTAS < 5 EM VERMELHO
    # ---------------------------------------------------------------------

    temp_out = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    df_final.to_excel(temp_out.name, index=False)

    # ---------------------------------------------------------
    #  GERAR CABEÇALHO DUPLO (MATÉRIA / BIMESTRE)
    # ---------------------------------------------------------

    def formatar_cabecalho_duplo(path):
        wb = load_workbook(path)
        ws = wb.active

        # Extrair matérias do df_final
        materias = {}
        for col in df_final.columns:
            if col == "ALUNO":
                continue
            nome, bi = col.split("_")  # Ciências_B1 → ["Ciências", "B1"]
            if nome not in materias:
                materias[nome] = []
            materias[nome].append(col)

        # Criar duas linhas extras no topo
        ws.insert_rows(1)
        ws.insert_rows(1)

        # Preencher nome "ALUNO"
        ws["A1"] = ""
        ws["A2"] = "ALUNO"

        col_excel = 2  # começa na coluna B

        for materia, colunas in materias.items():
            # Quantas colunas essa matéria ocupa?
            n_cols = len(colunas)

            # Cabeçalho nível 1 (mesclado)
            c1 = ws.cell(row=1, column=col_excel)
            c1.value = materia
            ws.merge_cells(start_row=1, start_column=col_excel,
                           end_row=1, end_column=col_excel + n_cols - 1)

            # Cabeçalho nível 2 (1ºBi, 2ºBi, 3ºBi)
            for i, colname in enumerate(colunas):
                bi = colname.split("_")[1].replace("B", "ºBi")
                ws.cell(row=2, column=col_excel + i, value=bi)

            col_excel += n_cols

        wb.save(path)

    # ---------------------------------------------------------
    # COLORIR NOTAS
    # ---------------------------------------------------------

    def colorir_notas(path):
        wb = load_workbook(path)
        ws = wb.active
        red = Font(color="FF0000", bold=True)

        for col in range(2, ws.max_column + 1):
            for row in range(3, ws.max_row + 1):  # notas começam na linha 3 agora
                val = ws.cell(row=row, column=col).value
                if isinstance(val, (int, float)) and val < 5:
                    ws.cell(row=row, column=col).font = red

        wb.save(path)

    # Aplicar cabeçalho e colorações
    formatar_cabecalho_duplo(temp_out.name)
    colorir_notas(temp_out.name)

    # Botão de download
    with open(temp_out.name, "rb") as f:
        st.download_button(
            "⬇️ Baixar Planilha Final Formatada",
            f.read(),
            file_name="notas_unificadas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

