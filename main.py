from src.app.interface import main

main()

# from dotenv import load_dotenv
# import os
# from googleapiclient.discovery import build
# from src.Auth.auth_sheets import autenticar
# from services.planilha_service import copiar_para_aba_existente

# load_dotenv()

# PLANILHA_ID_ORIGEM = os.getenv("PLANILHA_ID_ORIGEM")
# ABA_NOME_ORIGEM = os.getenv("ABA_NOME_ORIGEM")

# PLANILHA_ID_DESTINO_ASCGERAL = os.getenv("PLANILHA_ID_DESTINO_ASCGERAL")
# ABA_NOME_DESTINO_ASCGERAL = os.getenv("ABA_NOME_DESTINO_ASCGERAL")

# creds = autenticar()
# service = build('sheets', 'v4', credentials=creds)

# copiar_para_aba_existente(
#     service,
#     PLANILHA_ID_ORIGEM,
#     ABA_NOME_ORIGEM,
#     PLANILHA_ID_DESTINO_ASCGERAL,
#     ABA_NOME_DESTINO_ASCGERAL
# )

# print("Dados copiados com sucesso!")





