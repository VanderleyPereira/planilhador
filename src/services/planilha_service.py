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

def parse_colunas_protegidas(input_str):
    """Converte string 'A, C, F' para set de índices {0, 2, 5}"""
    if not input_str:
        return set()
    
    indices = set()
    partes = input_str.split(',')
    for p in partes:
        p = p.strip()
        if p:
            try:
                indices.add(col_para_idx(p))
            except ValueError:
                pass # Ignora colunas inválidas silenciosamente ou logar se necessário
    return indices

def calcular_intervalos_livres(indices_protegidos):
    """
    Retorna lista de tuplas (inicio, fim) representando colunas NÃO protegidas.
    'fim' pode ser None, indicando 'até o infinito'.
    Ex: protegidos={1} (B) -> [(0,0), (2, None)] -> A:A, C:Z
    """
    if not indices_protegidos:
        return [(0, None)]
    
    sorted_idx = sorted(list(indices_protegidos))
    intervalos = []
    
    current = 0
    for idx_proibido in sorted_idx:
        if idx_proibido > current:
            intervalos.append((current, idx_proibido - 1))
        current = idx_proibido + 1
    
    # Adiciona o intervalo final
    intervalos.append((current, None))
    
    return intervalos

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
        range=aba_origem,
        valueRenderOption='UNFORMATTED_VALUE',
        dateTimeRenderOption='SERIAL_NUMBER'
    ).execute()

    valores = result.get("values", [])

    if not valores:
        raise Exception("A aba de origem está vazia")

    # 2️⃣ Verificar e ajustar tamanho da aba destino (Auto-Resize)
    # Obtém metadados da planilha de destino para saber o SheetId e dimensões atuais
    destino_metadata = service.spreadsheets().get(spreadsheetId=planilha_destino_id).execute()
    aba_destino_info = next((s for s in destino_metadata['sheets'] if s['properties']['title'] == aba_destino), None)
    
    if not aba_destino_info:
        raise Exception(f"A aba '{aba_destino}' não foi encontrada na planilha de destino.")

    sheet_id = aba_destino_info['properties']['sheetId']
    grid_properties = aba_destino_info['properties'].get('gridProperties', {})
    linhas_atuais = grid_properties.get('rowCount', 0)
    colunas_atuais = grid_properties.get('columnCount', 0)

    linhas_necessarias = len(valores)
    # Descobre a maior largura de linha nos dados de origem
    colunas_necessarias = max(len(linha) for linha in valores) if valores else 0

    requests = []
    
    # Se precisar de mais linhas
    if linhas_necessarias > linhas_atuais:
        requests.append({
            "updateSheetProperties": {
                "properties": {
                    "sheetId": sheet_id,
                    "gridProperties": {"rowCount": linhas_necessarias}
                },
                "fields": "gridProperties.rowCount"
            }
        })
    
    # Se precisar de mais colunas
    if colunas_necessarias > colunas_atuais:
        requests.append({
            "updateSheetProperties": {
                "properties": {
                    "sheetId": sheet_id,
                    "gridProperties": {"columnCount": colunas_necessarias}
                },
                "fields": "gridProperties.columnCount"
            }
        })

    if requests:
        service.spreadsheets().batchUpdate(
            spreadsheetId=planilha_destino_id,
            body={"requests": requests}
        ).execute()

    # 3️⃣ Limpeza da aba destino (Batch Clear)
    # Parse das colunas protegidas
    indices_protegidos = parse_colunas_protegidas(coluna_protegida)
    intervalos_livres = calcular_intervalos_livres(indices_protegidos)
    
    ranges_to_clear = []
    for inicio, fim in intervalos_livres:
        col_start = idx_para_col(inicio)
        if fim is None:
            # Agora usamos o número real de colunas para evitar o ZZZ arbitrário
            col_end = idx_para_col(max(colunas_necessarias, colunas_atuais) - 1)
        else:
            col_end = idx_para_col(fim)
            
        ranges_to_clear.append(f"{aba_destino}!{col_start}:{col_end}")

    if ranges_to_clear:
        try:
            service.spreadsheets().values().batchClear(
                spreadsheetId=planilha_destino_id,
                body={"ranges": ranges_to_clear}
            ).execute()
        except Exception:
             for rng in ranges_to_clear:
                 try:
                     service.spreadsheets().values().clear(
                         spreadsheetId=planilha_destino_id, 
                         range=rng
                     ).execute()
                 except:
                     pass

    # 4️⃣ Colar dados na aba destino (Batch Update por blocos)
    total = len(valores)
    tamanho_bloco = 5000 

    if callback_progresso:
        callback_progresso(0)

    for i in range(0, total, tamanho_bloco):
        bloco = valores[i:i + tamanho_bloco]
        batch_data = []

        for inicio, fim in intervalos_livres:
            sub_bloco = []
            col_start_letter = idx_para_col(inicio)
            range_dest = f"{aba_destino}!{col_start_letter}{i+1}"
            
            has_data = False
            for linha in bloco:
                fatia = linha[inicio : (fim + 1 if fim is not None else None)]
                sub_bloco.append(fatia)
                if fatia: has_data = True
            
            if has_data:
                batch_data.append({
                    "range": range_dest,
                    "values": sub_bloco
                })

        if batch_data:
            try:
                service.spreadsheets().values().batchUpdate(
                    spreadsheetId=planilha_destino_id,
                    body={
                        "valueInputOption": "USER_ENTERED",
                        "data": batch_data
                    }
                ).execute()
            except Exception as e:
                raise Exception(f"Erro ao enviar lote de dados (linhas {i+1} a {i+len(bloco)}): {str(e)}")

        progresso = ((i + len(bloco)) / total) * 100
        if callback_progresso:
            callback_progresso(progresso)
        
        time.sleep(0.5)
