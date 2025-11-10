import pandas as pd

MAPEAMENTO_ABAS = {
    "1": {
        "titulo": "VERIFICAÇÃO E INSPEÇÃO MEC.",
        "descricao": "Análise mecânica e verificação de componentes",
        "colunas": ["Equipamento", "Quantidade", "Teste Realizado", "Observações / Justificativa"]
    },
    "2": {
        "titulo": "INSPEÇÃO VISUAL",
        "descricao": "Inspeção visual detalhada dos elementos",
        "colunas": ["SENSORES", "LOCAL INSTALADO", "TESTE REALIZADO", "OK", "NOK", "OBSERVAÇÕES"]
    },
    "3": {"titulo": "VALIDAÇÃO DE CIRCUITO", "descricao": "Verificação e validação dos circuitos elétricos", "colunas": ["Circuito", "Pontos de Teste", "Validação"]},
    "4": {"titulo": "ATERRAMENTO", "descricao": "Testes e verificação do sistema de aterramento", "colunas": ["Ponto", "Resistência", "Medição"]},
    "5": {"titulo": "DESEMPENHO DO SISTEMA", "descricao": "Avaliação do desempenho geral do sistema", "colunas": ["Sistema", "Parâmetro", "Medição"]},
    "6": {"titulo": "PROCEDIMENTO VERIFICAÇÃO CLP", "descricao": "Verificação dos procedimentos do CLP", "colunas": ["Tag", "Descrição", "Verificação"]}
}

def encontrar_cabecalho(df, aba_id=None):
    # Encontra as linhas de cabeçalho relevantes no DataFrame
    for idx, row in df.iterrows():
        if row.notna().sum() >= 3:  # Pelo menos 3 colunas não vazias
            valores = [str(val).strip().upper() for val in row if pd.notna(val)]
            
            if any(val in ["EQUIPAMENTO", "CIRCUITO", "PONTO", "SISTEMA", "TAG", "SENSOR"] for val in valores):
                return idx
    return 0

def carregar_abas(caminho_planilha):
    xlsx = pd.ExcelFile(caminho_planilha)
    abas_info = []
    
    for nome_aba in xlsx.sheet_names:
        for aba_id, info in MAPEAMENTO_ABAS.items():
            if info["titulo"] in nome_aba:
                df = pd.read_excel(caminho_planilha, sheet_name=nome_aba)
                cabecalho_idx = encontrar_cabecalho(df, aba_id)
                df = df.iloc[cabecalho_idx:]
                df = df.reset_index(drop=True)
                
                total_itens = len([row for idx, row in df.iterrows() if not pd.isna(row).all()])
                
                abas_info.append({
                    'id': aba_id,
                    'titulo': nome_aba,
                    'descricao': info["descricao"],
                    'total_itens': total_itens - 1
                })
                break
    
    return sorted(abas_info, key=lambda x: int(x['id']))

def carregar_itens(caminho_planilha, aba_id=None):
    xlsx = pd.ExcelFile(caminho_planilha)

    def processar_sheet(nome_aba, aba_id=None):
        df = pd.read_excel(caminho_planilha, sheet_name=nome_aba)
        cabecalho_idx = encontrar_cabecalho(df, aba_id)

        # ABA 2 - INSPEÇÃO VISUAL (FORÇAR NOVA ESTRUTURA)
        if aba_id == "2":
            # SEMPRE usar os novos cabeçalhos
            header_names = ["SENSORES", "LOCAL INSTALADO", "TESTE REALIZADO", "OK", "NOK", "OBSERVAÇÕES"]
            
            # Encontra a linha de dados (pula o cabeçalho original)
            if cabecalho_idx is not None and 0 <= cabecalho_idx < len(df):
                df_dados = df.iloc[cabecalho_idx + 1:].reset_index(drop=True)
            else:
                df_dados = df.reset_index(drop=True)

            # Remove linhas que contêm cabeçalhos repetidos
            df_dados = df_dados[
                ~df_dados.iloc[:, 0].astype(str).str.contains("EQUIPAMENTO", case=False, na=False)
            ]

            # Garante que temos pelo menos 6 colunas
            if len(df_dados.columns) < 6:
                # Adiciona colunas vazias se necessário
                for i in range(len(df_dados.columns), 6):
                    df_dados[f'extra_{i}'] = ""
            
            # Pega apenas as primeiras 6 colunas e aplica os NOVOS cabeçalhos
            df_dados = df_dados.iloc[:, :6]
            df_dados.columns = header_names

        # ABA 1 - VERIFICAÇÃO E INSPEÇÃO MEC.
        elif aba_id == "1":
            header_names = ["Equipamento", "Quantidade", "Teste Realizado", "OK", "NOK", "Observações / Justificativa"]

            if cabecalho_idx is not None and 0 <= cabecalho_idx < len(df):
                df_dados = df.iloc[cabecalho_idx + 1:].reset_index(drop=True)
            else:
                df_dados = df.reset_index(drop=True)

            df_dados = df_dados[
                ~df_dados.iloc[:, 0].astype(str).str.contains("EQUIPAMENTO", case=False, na=False)
            ]

            df_dados = df_dados.iloc[:, :3]
            df_dados.columns = ["Equipamento", "Quantidade", "Teste Realizado"]

        # OUTRAS ABAS
        else:
            if cabecalho_idx is not None and 0 <= cabecalho_idx < len(df):
                header_row = df.iloc[cabecalho_idx]
                header_names = [str(x).strip() for x in header_row.tolist() if pd.notna(x)]
                df_dados = df.iloc[cabecalho_idx + 1:].reset_index(drop=True)
            else:
                header_names = [str(c) for c in df.columns[:6]]
                df_dados = df.reset_index(drop=True)

            if len(df_dados) > 0:
                df_dados = df_dados.iloc[1:].reset_index(drop=True)

            # Garante que temos cabeçalhos para todas as colunas
            while len(header_names) < len(df_dados.columns):
                header_names.append(f'Coluna {len(header_names) + 1}')
            
            df_dados.columns = header_names[:len(df_dados.columns)]

        # Processa os itens
        itens = []
        for idx, row in df_dados.iterrows():
            if pd.isna(row).all():
                continue

            item = {}
            valores_validos = 0
            
            for col_idx, col_name in enumerate(df_dados.columns[:6]):
                val = row[col_name] if col_name in row.index else ""
                if pd.notna(val) and str(val).strip():
                    item[f'coluna_{col_idx + 1}'] = str(val).strip()
                    valores_validos += 1
                else:
                    item[f'coluna_{col_idx + 1}'] = ""

            if valores_validos >= 1:
                item['aba'] = nome_aba
                itens.append(item)

        show_quantity_test = (aba_id == "1")
        return itens, header_names, show_quantity_test

    # Processamento da aba solicitada
    if aba_id:
        info = MAPEAMENTO_ABAS.get(str(aba_id))
        if not info:
            return {'items': [], 'headers': []}

        target = info['titulo']
        sheet_name = None
        for nome in xlsx.sheet_names:
            if target in nome:
                sheet_name = nome
                break

        if not sheet_name:
            return {'items': [], 'headers': [], 'show_quantity_test': False}

        itens, headers, show_quantity_test = processar_sheet(sheet_name, aba_id)
        
        # PARA A ABA 2, SEMPRE retornar os cabeçalhos novos
        if aba_id == "2":
            headers = ["SENSORES", "LOCAL INSTALADO", "TESTE REALIZADO", "OK", "NOK", "OBSERVAÇÕES"]
            
        return {'items': itens, 'headers': headers, 'show_quantity_test': show_quantity_test}

    todas_abas = []
    for nome_aba in xlsx.sheet_names:
        for aid, info in MAPEAMENTO_ABAS.items():
            if info["titulo"] in nome_aba:
                itens, _, _ = processar_sheet(nome_aba, aid)
                if itens:
                    todas_abas.extend(itens)
                break

    return {'items': todas_abas, 'headers': [], 'show_quantity_test': False}

def salvar_resultados(caminho_saida, dados):
    df = pd.DataFrame(dados)
    df.to_excel(caminho_saida, index=False)