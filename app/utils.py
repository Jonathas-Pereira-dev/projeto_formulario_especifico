import pandas as pd

MAPEAMENTO_ABAS = {
    "1": {"titulo": "VERIFICAÇÃO E INSPEÇÃO MEC.", "descricao": "Análise mecânica e verificação de componentes"},
    "2": {"titulo": "INSPEÇÃO VISUAL", "descricao": "Inspeção visual detalhada dos elementos", "colunas": ["Equipamento", "Quantidade", "Inspeção"]},
    "3": {"titulo": "VALIDAÇÃO DE CIRCUITO", "descricao": "Verificação e validação dos circuitos elétricos", "colunas": ["Circuito", "Pontos de Teste", "Validação"]},
    "4": {"titulo": "ATERRAMENTO", "descricao": "Testes e verificação do sistema de aterramento", "colunas": ["Ponto", "Resistência", "Medição"]},
    "5": {"titulo": "DESEMPENHO DO SISTEMA", "descricao": "Avaliação do desempenho geral do sistema", "colunas": ["Sistema", "Parâmetro", "Medição"]},
    "6": {"titulo": "PROCEDIMENTO VERIFICAÇÃO CLP", "descricao": "Verificação dos procedimentos do CLP", "colunas": ["Tag", "Descrição", "Verificação"]}
}

def encontrar_cabecalho(df):
    """Encontra as linhas de cabeçalho relevantes no DataFrame."""
    for idx, row in df.iterrows():
        if row.notna().sum() >= 3:  # Pelo menos 3 colunas não vazias
            valores = [str(val).strip() for val in row if pd.notna(val)]
            if any(val in ["Equipamento", "Circuito", "Ponto", "Sistema", "Tag"] for val in valores):
                return idx
    return 0

def carregar_abas(caminho_planilha):
    xlsx = pd.ExcelFile(caminho_planilha)
    abas_info = []
    
    for nome_aba in xlsx.sheet_names:
        for aba_id, info in MAPEAMENTO_ABAS.items():
            if info["titulo"] in nome_aba:
                df = pd.read_excel(caminho_planilha, sheet_name=nome_aba)
                cabecalho_idx = encontrar_cabecalho(df)
                df = df.iloc[cabecalho_idx:]
                df = df.reset_index(drop=True)
                
                total_itens = len([row for idx, row in df.iterrows() if not pd.isna(row).all()])
                
                abas_info.append({
                    'id': aba_id,
                    'titulo': nome_aba,
                    'descricao': info["descricao"],
                    'total_itens': total_itens - 1  # Subtrai a linha de cabeçalho
                })
                break
    
    return sorted(abas_info, key=lambda x: int(x['id']))

def carregar_itens(caminho_planilha, aba_id=None):
    xlsx = pd.ExcelFile(caminho_planilha)

    def processar_sheet(nome_aba):
        df = pd.read_excel(caminho_planilha, sheet_name=nome_aba)
        cabecalho_idx = encontrar_cabecalho(df)

        header_names = []
        if cabecalho_idx is not None and cabecalho_idx >= 0 and cabecalho_idx < len(df):
            header_row = df.iloc[cabecalho_idx]
            # Extrai nomes de cabeçalho das células identificadas
            header_names = [str(x).strip() for x in header_row.tolist() if pd.notna(x)]
            # os dados reais começam na linha seguinte
            df_dados = df.iloc[cabecalho_idx + 1 :].reset_index(drop=True)
        else:
            # Usa o cabeçalho do DataFrame (caso as colunas já estejam nomeadas)
            header_names = [str(c) for c in df.columns[:6]]
            df_dados = df.reset_index(drop=True)

        itens = []

        # Remove a primeira linha de dados conforme solicitado (descarta linha de títulos/exemplo)
        if len(df_dados) > 0:
            df_dados = df_dados.iloc[1:].reset_index(drop=True)

        for idx, row in df_dados.iterrows():
            if pd.isna(row).all():
                continue

            item = {}
            valores_validos = 0

            # Iterate over up to 6 columns (preservando ordem)
            for col_idx, col in enumerate(df_dados.columns[:6]):
                val = row[col]
                if pd.notna(val):
                    valor = str(val).strip()
                    if valor:
                        item[f'coluna_{col_idx + 1}'] = valor
                        valores_validos += 1

            # Considera como item válido se houver pelo menos 1 valor (mais permissivo)
            if valores_validos >= 1:
                item['aba'] = nome_aba
                itens.append(item)

        # Garante que header_names tenha pelo menos o mesmo número de colunas consideradas
        if not header_names:
            header_names = [f'Coluna {i+1}' for i in range(6)]

        # Decidir se mostramos as colunas Quantidade / Teste Realizado (usuario pediu ocultar)
        # Por enquanto, escondemos essas colunas globalmente conforme instruído
        show_quantity_test = False

        return itens, header_names, show_quantity_test

    # Se um aba_id foi solicitado, processa apenas a folha correspondente
    if aba_id:
        info = MAPEAMENTO_ABAS.get(str(aba_id))
        if not info:
            return {'items': [], 'headers': []}

        # Procura o nome da folha que contém o título configurado
        target = info['titulo']
        sheet_name = None
        for nome in xlsx.sheet_names:
            if target in nome:
                sheet_name = nome
                break

        if not sheet_name:
            return {'items': [], 'headers': [], 'show_quantity_test': False}

        itens, headers, show_quantity_test = processar_sheet(sheet_name)
        return {'items': itens, 'headers': headers, 'show_quantity_test': show_quantity_test}

    # Caso contrário, processa todas as folhas mapeadas
    todas_abas = []
    for nome_aba in xlsx.sheet_names:
        for aid, info in MAPEAMENTO_ABAS.items():
            if info['titulo'] in nome_aba:
                itens, _, _ = processar_sheet(nome_aba)
                if itens:
                    todas_abas.extend(itens)
                break

    return {'items': todas_abas, 'headers': [], 'show_quantity_test': False}

def salvar_resultados(caminho_saida, dados):
    df = pd.DataFrame(dados)
    df.to_excel(caminho_saida, index=False)
