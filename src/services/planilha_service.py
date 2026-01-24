import time

def copiar_para_aba_existente(
    service,
    planilha_origem_id,
    aba_origem,
    planilha_destino_id,
    aba_destino,
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

    # 2️⃣ Limpar aba destino
    service.spreadsheets().values().clear(
        spreadsheetId=planilha_destino_id,
        range=aba_destino
    ).execute()

    # 3️⃣ Colar dados na aba destino (em blocos)
    total = len(valores)
    tamanho_bloco = 5000  # Blocos de 2000 linhas para envio fracionado

    if callback_progresso:
        callback_progresso(0)

    for i in range(0, total, tamanho_bloco):
        bloco = valores[i:i + tamanho_bloco]
        
        # Define a célula inicial do bloco (Ex: A1, A2001, A4001...)
        range_bloco = f"{aba_destino}!A{i+1}"

        try:
            service.spreadsheets().values().update(
                spreadsheetId=planilha_destino_id,
                range=range_bloco,
                valueInputOption="RAW",
                body={"values": bloco}
            ).execute()
        except Exception as e:
            raise Exception(f"Erro ao enviar bloco {i} até {i+len(bloco)}: {str(e)}")

        # Calcula progresso
        progresso = ((i + len(bloco)) / total) * 100
        
        if callback_progresso:
            callback_progresso(progresso)
        
        # Pausa para evitar Rate Limit
        time.sleep(0.5)
