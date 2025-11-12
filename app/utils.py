import pandas as pd

MAPEAMENTO_ABAS = {
    "1": {
        "titulo": "VERIFICAÇÃO E INSPEÇÃO MEC.",
        "descricao": "Análise mecânica e verificação de componentes",
        "colunas": ["Equipamento", "Quantidade", "Teste Realizado", "OK", "NOK", "Observações / Justificativa"]
    },
    "2": {
        "titulo": "INSPEÇÃO VISUAL",
        "descricao": "Inspeção visual detalhada dos elementos",
        "colunas": ["SENSORES", "LOCAL INSTALADO", "TESTE REALIZADO", "OK", "NOK", "OBSERVAÇÕES"]
    },
    "3": {
        "titulo": "VALIDAÇÃO DE CIRCUITO",
        "descricao": "Verificação e validação dos circuitos elétricos",
        "colunas": ["EQUIPAMENTO", "PONTO 1", "TAG P1", "PONTO 2", "TAG P2", "OK", "NOK", "OBSERVAÇÕES"]
    },
    "4": {
        "titulo": "ATERRAMENTO",
        "descricao": "Testes e verificação do sistema de aterramento",
        "colunas": ["PONTO DE ATERRAMENTO", "OK", "NOK", "OBSERVAÇÕES"]
    },
    "5": {
        "titulo": "DESEMPENHO DO SISTEMA",
        "descricao": "Avaliação do desempenho geral do sistema",
        "colunas": [
            "EQUIPAMENTO",
            "PONTOS ALIMENTAÇÃO / ATERRAMENTO",
            "ALIMENTAÇÃO TEÓRICA",
            "ALIMENTAÇÃO AFERIDA",
            "OK",
            "NOK",
            "OBSERVAÇÕES"
        ]
    },
    "6": {
        "titulo": "PROCEDIMENTO VERIFICAÇÃO CLP",
        "descricao": "Verificação dos procedimentos do CLP",
        "colunas": ["EQUIPAMENTO", "OK", "NOK", "OBSERVAÇÕES"]
    }
}

# Mapeamento para formulário campo
MAPEAMENTO_FORMULARIO_CAMPO = {
    "comunicacao": {
        "aba": "TESTE DE COMUNICAÇÃO ENTRE CLP",
        "colunas": ["EQUIPAMENTO", "STATUS DO PAINEL", "ITEM DO PT", "OK", "NOK", "OBSERVAÇÕES"]
    },
    "sensores_digitais": {
        "aba": "TESTES SENSORES DIGITAIS", 
        "colunas": ["EQUIPAMENTO", "SENSOR", "ITEM DO PT", "ESTADO", "OK", "NOK", "OBSERVAÇÃO"]
    },
    "sensores_analogicos": {
        "aba": "SENSORES ANALÓGICOS",
        "colunas": ["EQUIPAMENTO", "SENSOR", "ITEM DO PT", "VALOR PREVISTO", "VARIÁVEIS MEDIDAS NO CLP", "VARIÁVEIS MEDIDAS NO EPM", "OK", "NOK", "OBSERVAÇÕES"]
    }
}


def encontrar_cabecalho(df, aba_id=None):
    for idx, row in df.iterrows():
        if row.notna().sum() >= 3:
            valores = [str(val).strip().upper() for val in row if pd.notna(val)]
            if any(val in ["EQUIPAMENTO", "CIRCUITO", "PONTO", "SISTEMA", "TAG", "SENSOR", "SENSORES", "ATERRAMENTO"]
                   for val in valores):
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
                df = df.iloc[cabecalho_idx:].reset_index(drop=True)

                total_itens = len([row for _, row in df.iterrows() if not pd.isna(row).all()])

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

        if aba_id in MAPEAMENTO_ABAS:
            header_names = MAPEAMENTO_ABAS[aba_id]["colunas"]
        else:
            header_names = [str(c) for c in df.columns[:6]]

        if cabecalho_idx is not None and 0 <= cabecalho_idx < len(df):
            df_dados = df.iloc[cabecalho_idx + 1:].reset_index(drop=True)
        else:
            df_dados = df.reset_index(drop=True)

        primeira_coluna = df_dados.columns[0]
        df_dados = df_dados[
            ~df_dados[primeira_coluna].astype(str).str.contains(
                "EQUIPAMENTO|SENSOR|SENSORES|ATERRAMENTO|ALIMENTAÇÃO", case=False, na=False)
        ]

        while len(df_dados.columns) < len(header_names):
            df_dados[f"extra_{len(df_dados.columns)}"] = ""

        df_dados = df_dados.iloc[:, :len(header_names)]
        df_dados.columns = header_names

        itens = []
        for _, row in df_dados.iterrows():
            if pd.isna(row).all():
                continue
            item = {}
            for i, col in enumerate(df_dados.columns[:8]):
                val = row[col]
                item[f"coluna_{i + 1}"] = str(val).strip() if pd.notna(val) else ""
            item["aba"] = nome_aba
            itens.append(item)

        show_quantity_test = (aba_id == "1")
        return itens, header_names, show_quantity_test

    if aba_id:
        info = MAPEAMENTO_ABAS.get(str(aba_id))
        if not info:
            return {'items': [], 'headers': [], 'show_quantity_test': False}

        target = info["titulo"]
        sheet_name = next((n for n in xlsx.sheet_names if target in n), None)

        if not sheet_name:
            return {'items': [], 'headers': [], 'show_quantity_test': False}

        itens, headers, show_quantity_test = processar_sheet(sheet_name, aba_id)
        return {'items': itens, 'headers': headers, 'show_quantity_test': show_quantity_test}

    todas_abas = []
    for nome_aba in xlsx.sheet_names:
        for aid, info in MAPEAMENTO_ABAS.items():
            if info["titulo"] in nome_aba:
                itens, _, _ = processar_sheet(nome_aba, aid)
                todas_abas.extend(itens)
                break

    return {'items': todas_abas, 'headers': [], 'show_quantity_test': False}


def carregar_formulario_campo(caminho_planilha, estacao):
    """
    Carrega dados do formulário campo filtrando por estação
    """
    try:
        dados = {}
        
        for tipo, info in MAPEAMENTO_FORMULARIO_CAMPO.items():
            df = pd.read_excel(caminho_planilha, sheet_name=info["aba"])
            
            # Encontrar cabeçalho
            cabecalho_idx = encontrar_cabecalho(df)
            
            if cabecalho_idx is not None and 0 <= cabecalho_idx < len(df):
                df_dados = df.iloc[cabecalho_idx + 1:].reset_index(drop=True)
            else:
                df_dados = df.reset_index(drop=True)
            
            # Definir colunas
            colunas_disponiveis = len(df_dados.columns)
            colunas_necessarias = len(info["colunas"])
            
            if colunas_disponiveis < colunas_necessarias:
                # Adicionar colunas extras se necessário
                for i in range(colunas_disponiveis, colunas_necessarias):
                    df_dados[f"extra_{i}"] = ""
            
            df_dados = df_dados.iloc[:, :colunas_necessarias]
            df_dados.columns = info["colunas"]
            
            # Filtrar por estação se a coluna existir
            if 'ESTAÇÃO' in df_dados.columns:
                df_filtrado = df_dados[df_dados['ESTAÇÃO'].astype(str).str.upper() == estacao.upper()]
                df_filtrado = df_filtrado.drop('ESTAÇÃO', axis=1)
            else:
                df_filtrado = df_dados
            
            # Converter para lista de dicionários
            itens = []
            for _, row in df_filtrado.iterrows():
                if pd.isna(row).all():
                    continue
                item = {}
                for col in df_filtrado.columns:
                    val = row[col]
                    item[col] = str(val).strip() if pd.notna(val) else ""
                itens.append(item)
            
            dados[tipo] = itens
        
        return dados
        
    except Exception as e:
        print(f"Erro ao carregar formulário campo: {e}")
        return {}


def salvar_resultados(caminho_saida, dados):
    df = pd.DataFrame(dados)
    df.to_excel(caminho_saida, index=False)