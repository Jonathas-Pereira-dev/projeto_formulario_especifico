import pandas as pd

def carregar_itens(caminho_planilha):
    # Lê todas as abas da planilha
    xlsx = pd.ExcelFile(caminho_planilha)
    todas_abas = []
    
    for nome_aba in xlsx.sheet_names:
        # Lê cada aba da planilha
        df = pd.read_excel(caminho_planilha, sheet_name=nome_aba)
        
        # Procura pelas colunas relevantes
        equipamentos = []
        for idx, row in df.iterrows():
            # Pula linhas vazias
            if pd.isna(row).all():
                continue
                
            # Procura por linhas que contenham equipamentos
            equipamento = None
            quantidade = None
            atividade = None
            
            for col in df.columns:
                val = row[col]
                if pd.notna(val):
                    # Procura pela coluna de equipamentos
                    if isinstance(val, str):
                        if not equipamento:
                            equipamento = str(val).strip()
                        # Procura pela quantidade
                        elif not quantidade and str(val).strip().endswith(('Peça', 'Peças', 'Peca', 'Pecas')):
                            quantidade = str(val).strip()
                        # Procura pela atividade
                        elif not atividade and 'Análise' in str(val):
                            atividade = str(val).strip()
            
            if equipamento and quantidade and atividade:
                equipamentos.append({
                    'aba': nome_aba,
                    'equipamento': equipamento,
                    'quantidade': quantidade,
                    'atividade': atividade,
                })
        
        if equipamentos:
            todas_abas.extend(equipamentos)
    
    return todas_abas

def salvar_resultados(caminho_saida, dados):
    df = pd.DataFrame(dados)
    df.to_excel(caminho_saida, index=False)
