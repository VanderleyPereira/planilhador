import time

def col_para_idx(col):
    """Converte letra da coluna (A, B, CU) para índice (0, 1, 98)"""
    col = col.upper().strip()
    if not col.isalpha():
        raise ValueError(f"Coluna inválida: {col}")
    idx = 0
    for char in col:
        idx = idx * 26 + (ord(char) - ord('A') + 1)
    return idx - 1

def idx_para_col(idx):
    """Converte índice (0, 1, 98) para letra da coluna (A, B, CU)"""
    idx += 1
    letras = ""
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        letras = chr(65 + rem) + letras
    return letras

def copiar_para_aba_existente(
    service,
    planilha_origem_id,
    aba_origem,
    planilha_destino_id,
    aba_destino,
    coluna_protegida=None,
    callback_progresso=None
):
    # 1️⃣ Ler dados da origem
    result = service.spreadsheets().values().get(
        spreadsheetId=planilha_origem_id,
        range=aba_origem
    ).execute()

    valores = result.get("values", [])

    if not valores:
        raise Exception("A aba de origem está vazia")

    # 2️⃣ Limpeza da aba destino
    if not coluna_protegida:
        # Se não houver proteção, limpa tudo
        service.spreadsheets().values().clear(
            spreadsheetId=planilha_destino_id,
            range=aba_destino
        ).execute()
    else:
        # Se houver proteção (ex: D), limpa em duas partes
        idx_prot = col_para_idx(coluna_protegida)
        
        # Parte 1: Antes da protegida (Ex: de A até C se a protegida for D)
        if idx_prot > 0:
            letra_fim_p1 = idx_para_col(idx_prot - 1)
            range_p1 = f"{aba_destino}!A:{letra_fim_p1}"
            service.spreadsheets().values().clear(spreadsheetId=planilha_destino_id, range=range_p1).execute()
        
        # Parte 2: Depois da protegida (Ex: de E em diante se a protegida for D)
        try:
            letra_ini_p2 = idx_para_col(idx_prot + 1)
            range_p2 = f"{aba_destino}!{letra_ini_p2}:ZZZ"
            service.spreadsheets().values().clear(spreadsheetId=planilha_destino_id, range=range_p2).execute()
        except Exception:
            # Se a coluna protegida for a última ou estiver além do grid, ignoramos a limpeza do resto
            pass

    # 3️⃣ Colar dados na aba destino (em blocos)
    total = len(valores)
    tamanho_bloco = 5000  # Blocos de 2000 linhas para envio fracionado

    if callback_progresso:
        callback_progresso(0)

    for i in range(0, total, tamanho_bloco):
        bloco = valores[i:i + tamanho_bloco]
        
        if not coluna_protegida:
            # Lógica Normal: Cola bloco inteiro começando em A
            range_bloco = f"{aba_destino}!A{i+1}"
            try:
                service.spreadsheets().values().update(
                    spreadsheetId=planilha_destino_id,
                    range=range_bloco,
                    valueInputOption="RAW",
                    body={"values": bloco}
                ).execute()
            except Exception as e:
                raise Exception(f"Erro ao enviar bloco (normal) {i+1}: {str(e)}")
        else:
            # Lógica com Coluna Protegida
            idx_prot = col_para_idx(coluna_protegida)
            
            # Fatiar bloco em duas partes (esquerda e direita da coluna protegida)
            bloco_esq = []
            bloco_dir = []
            
            for linha in bloco:
                esq = linha[:idx_prot]
                dir = linha[idx_prot + 1:] if len(linha) > idx_prot else []
                bloco_esq.append(esq)
                bloco_dir.append(dir)
            
            # Envia Parte Esquerda (se existir e não for apenas listas vazias)
            if bloco_esq and any(linha for linha in bloco_esq if linha):
                range_esq = f"{aba_destino}!A{i+1}"
                try:
                    service.spreadsheets().values().update(
                        spreadsheetId=planilha_destino_id,
                        range=range_esq,
                        valueInputOption="RAW",
                        body={"values": bloco_esq}
                    ).execute()
                except Exception as e:
                    raise Exception(f"Erro ao enviar parte esquerda no bloco {i+1}: {str(e)}")
            
            # Envia Parte Direita (se existir e não for apenas listas vazias)
            if bloco_dir and any(linha for linha in bloco_dir if linha):
                letra_ini_p2 = idx_para_col(idx_prot + 1)
                range_dir = f"{aba_destino}!{letra_ini_p2}{i+1}"
                try:
                    service.spreadsheets().values().update(
                        spreadsheetId=planilha_destino_id,
                        range=range_dir,
                        valueInputOption="RAW",
                        body={"values": bloco_dir}
                    ).execute()
                except Exception as e:
                    raise Exception(f"Erro ao enviar parte direita (após {coluna_protegida}) no bloco {i+1}: {str(e)}")

        # Calcula progresso
        progresso = ((i + len(bloco)) / total) * 100
        
        if callback_progresso:
            callback_progresso(progresso)
        
        # Pausa para evitar Rate Limit
        time.sleep(0.5)
