import pandas as pd
import os

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

# Mapeamento para formulário campo - VERIFICADO E CORRETO
MAPEAMENTO_FORMULARIO_CAMPO = {
    "comunicacao": {
        "aba": "TESTE DE COMUNICAÇÃO ENTRE CLP",
        "colunas": ["ESTAÇÃO", "EQUIPAMENTO", "STATUS DO PAINEL", "ITEM DO PT", "OK", "NOK", "OBSERVAÇÕES"]
    },
    "sensores_digitais": {
        "aba": "TESTES SENSORES DIGITAIS",
        "colunas": ["ESTAÇÃO", "EQUIPAMENTO", "SENSOR", "ITEM DO PT", "ESTADO", "OK", "NOK", "OBSERVAÇÃO"]
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
            if any(val in ["EQUIPAMENTO", "CIRCUITO", "PONTO", "SISTEMA", "TAG", "SENSOR", "SENSORES", "ATERRAMENTO", "ESTAÇÃO", "ESTACAO"]
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
        
        print(f"🎯 INICIANDO CARREGAMENTO PARA ESTAÇÃO: {estacao}")
        print(f"📁 Planilha: {caminho_planilha}")
        
        if not os.path.exists(caminho_planilha):
            print(f"❌ ARQUIVO NÃO ENCONTRADO: {caminho_planilha}")
            return {}
        
        xlsx = pd.ExcelFile(caminho_planilha)
        print(f"📑 Abas disponíveis: {xlsx.sheet_names}")
        
        for tipo, info in MAPEAMENTO_FORMULARIO_CAMPO.items():
            try:
                print(f"\n🔍 === PROCESSANDO ABA: {info['aba']} ===")
                
                if info['aba'] not in xlsx.sheet_names:
                    print(f"❌ ABA NÃO ENCONTRADA: {info['aba']}")
                    dados[tipo] = []
                    continue
                
                df = pd.read_excel(caminho_planilha, sheet_name=info["aba"])
                print(f"✅ Aba carregada - {len(df)} linhas")
                
                # ESTRATÉGIA ESPECÍFICA PARA CADA ABA
                if tipo == "comunicacao":
                    print("🎯 ESTRATÉGIA PARA COMUNICAÇÃO")
                    # Linha 1 é o cabeçalho (ESTAÇÃO, EQUIPAMENTO, etc.)
                    df.columns = df.iloc[1]  # Usar linha 1 como cabeçalho
                    df = df.iloc[2:].reset_index(drop=True)  # Dados começam na linha 2
                    
                elif tipo == "sensores_digitais":
                    print("🎯 ESTRATÉGIA PARA SENSORES DIGITAIS")
                    # Procurar linha com "ESTAÇÃO" e "EQUIPAMENTO"
                    cabecalho_idx = None
                    for idx, row in df.iterrows():
                        linha_str = ' '.join([str(cell).upper() for cell in row if pd.notna(cell)])
                        if "ESTAÇÃO" in linha_str and "EQUIPAMENTO" in linha_str:
                            cabecalho_idx = idx
                            break
                    
                    if cabecalho_idx is not None:
                        df.columns = df.iloc[cabecalho_idx]
                        df = df.iloc[cabecalho_idx + 1:].reset_index(drop=True)
                    else:
                        df.columns = df.iloc[0]
                        df = df.iloc[1:].reset_index(drop=True)
                        
                else:  # sensores_analogicos
                    print("🎯 ESTRATÉGIA PARA SENSORES ANALÓGICOS")
                    # Procurar linha com "EQUIPAMENTO" e "SENSOR"
                    cabecalho_idx = None
                    for idx, row in df.iterrows():
                        linha_str = ' '.join([str(cell).upper() for cell in row if pd.notna(cell)])
                        if "EQUIPAMENTO" in linha_str and "SENSOR" in linha_str:
                            cabecalho_idx = idx
                            break
                    
                    if cabecalho_idx is not None:
                        df.columns = df.iloc[cabecalho_idx]
                        df = df.iloc[cabecalho_idx + 1:].reset_index(drop=True)
                    else:
                        df.columns = df.iloc[0]
                        df = df.iloc[1:].reset_index(drop=True)
                
                # Remover linhas vazias
                df = df.dropna(how='all')
                print(f"📈 Dados processados: {len(df)} linhas")
                
                # FILTRAGEM POR ESTAÇÃO
                if tipo in ["comunicacao", "sensores_digitais"]:
                    # Encontrar coluna ESTAÇÃO
                    coluna_estacao = None
                    for col in df.columns:
                        if str(col).upper().strip() in ['ESTAÇÃO', 'ESTACAO', 'ESTAÇAO']:
                            coluna_estacao = col
                            break
                    
                    if coluna_estacao:
                        print(f"📍 Coluna de estação: '{coluna_estacao}'")
                        # Padronizar valores
                        df[coluna_estacao] = df[coluna_estacao].astype(str).str.upper().str.strip()
                        valores_unicos = df[coluna_estacao].unique()
                        print(f"📋 Valores únicos: {list(valores_unicos)}")
                        
                        # Filtrar por estação
                        df_filtrado = df[df[coluna_estacao] == estacao.upper()]
                        print(f"✅ {len(df_filtrado)} linhas para estação {estacao}")
                        
                        # Remover coluna ESTAÇÃO
                        df_filtrado = df_filtrado.drop(coluna_estacao, axis=1)
                    else:
                        print("❌ Coluna ESTAÇÃO não encontrada")
                        df_filtrado = df
                else:
                    # Sensores analógicos não tem filtro por estação
                    df_filtrado = df
                    print("ℹ️ Sensores analógicos - sem filtro por estação")
                
                # PREPARAR COLUNAS FINAIS
                colunas_necessarias = info["colunas"]
                
                # Adicionar colunas faltantes
                for col in colunas_necessarias:
                    if col not in df_filtrado.columns:
                        df_filtrado[col] = ""
                        print(f"➕ Adicionada coluna: {col}")
                
                # Selecionar e ordenar colunas
                df_filtrado = df_filtrado[colunas_necessarias]
                
                # CONVERTER PARA DICIONÁRIOS
                itens = []
                for _, row in df_filtrado.iterrows():
                    if pd.isna(row).all():
                        continue
                    
                    item = {}
                    for col in df_filtrado.columns:
                        val = row[col]
                        item[col] = str(val).strip() if pd.notna(val) else ""
                    
                    # Verificar se tem dados válidos
                    if any(value.strip() for value in item.values() if value):
                        itens.append(item)
                
                print(f"🎉 {len(itens)} itens carregados")
                if itens:
                    print(f"📄 Primeiro item: {itens[0]}")
                dados[tipo] = itens
                
            except Exception as e:
                print(f"❌ Erro na aba {info['aba']}: {e}")
                import traceback
                traceback.print_exc()
                dados[tipo] = []
        
        print(f"\n📊 RESUMO FINAL:")
        print(f"   Comunicação: {len(dados.get('comunicacao', []))} itens")
        print(f"   Sensores Digitais: {len(dados.get('sensores_digitais', []))} itens")
        print(f"   Sensores Analógicos: {len(dados.get('sensores_analogicos', []))} itens")
        
        return dados
        
    except Exception as e:
        print(f"🚨 ERRO GERAL: {e}")
        import traceback
        traceback.print_exc()
        return {}

def salvar_resultados(caminho_saida, dados):
    df = pd.DataFrame(dados)
    df.to_excel(caminho_saida, index=False)