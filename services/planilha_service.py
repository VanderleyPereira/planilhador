def copiar_para_aba_existente(
    service,
    planilha_origem_id,
    aba_origem,
    planilha_destino_id,
    aba_destino
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

    # 3️⃣ Colar dados na aba destino
    service.spreadsheets().values().update(
        spreadsheetId=planilha_destino_id,
        range=aba_destino,
        valueInputOption="RAW",
        body={"values": valores}
    ).execute()




















    

# import time

# def copiar_para_aba_existente(
#     service,
#     planilha_origem_id,
#     aba_origem,
#     planilha_destino_id,
#     aba_destino
# ):
#     # 1️⃣ Ler dados da origem
#     result = service.spreadsheets().values().get(
#         spreadsheetId=planilha_origem_id,
#         range=aba_origem
#     ).execute()

#     valores = result.get("values", [])

#     if not valores:
#         raise Exception("A aba de origem está vazia")

#     # 2️⃣ Limpar aba destino
#     service.spreadsheets().values().clear(
#         spreadsheetId=planilha_destino_id,
#         range=aba_destino
#     ).execute()

#     # 3️⃣ Colar dados na aba destino
#     # 3️⃣ Colar dados na aba destino (em blocos)
#     total = len(valores)
#     tamanho_bloco = 5000  # Aumentado para 2500 para otimizar 25k linhas (aprox. 10 chamadas)

#     print(f"Iniciando cópia de {total} linhas em blocos de {tamanho_bloco}...")

#     for i in range(0, total, tamanho_bloco):
#         bloco = valores[i:i + tamanho_bloco]
        
#         # Define a célula inicial do bloco (Ex: A1, A2501, A5001...)
#         range_bloco = f"{aba_destino}!A{i+1}"

#         try:
#             service.spreadsheets().values().update(
#                 spreadsheetId=planilha_destino_id,
#                 range=range_bloco,
#                 valueInputOption="RAW",
#                 body={"values": bloco}
#             ).execute()
#         except Exception as e:
#             print(f"❌ ERRO ao enviar bloco {i} até {i+len(bloco)}: {e}")
#             # Opcional: tentar novamente ou parar o script
#             raise e

#         progresso = ((i + len(bloco)) / total) * 100
#         print(f"Progresso: {progresso:.2f}%")
        
#         # Pausa de segurança para evitar limites de API (Rate Limiting)
#         time.sleep(1)
