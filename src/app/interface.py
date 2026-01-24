import os
import json
import threading
import FreeSimpleGUI as sg
from googleapiclient.discovery import build
from src.auth.auth_sheets import autenticar
from src.services.planilha_service import copiar_para_aba_existente
from src.icon.icon import ICONE_BASE64

# Arquivo de configuração
CONFIG_FILE = "planilhas_config.json"

def load_config():
    """Carrega configurações do arquivo JSON"""
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            return {"planilhas": []}
    return {"planilhas": []}

def save_config(config):
    """Salva configurações no arquivo JSON"""
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(config, f, indent=4, ensure_ascii=False)

def criar_modal_config_origem():
    """Cria modal para configurar nova origem"""
    sg.theme('DarkGrey14')
    
    # Cores (Tema Cinza Profissional)
    COLOR_BOTON_START= '#E0E0E0'
    BG_COLOR = '#2D2D2D'   # Cinza Fundo Principal
    INPUT_BG = '#383838'   # Cinza Inputs/Frames
    TEXT_COLOR = '#E0E0E0' # Texto Claro
    BTN_SUCCESS = '#4CAF50' # Verde Material
    BTN_DANGER = '#E57373'  # Vermelho Suave
    ACCENT_COLOR = '#64B5F6' # Azul Claro para destaque
    
    sg.set_options(background_color=BG_COLOR, text_element_background_color=BG_COLOR, element_background_color=BG_COLOR, input_elements_background_color=INPUT_BG, input_text_color=TEXT_COLOR, text_color=TEXT_COLOR)
    
    layout = [
        [sg.Text("⚙️ Planilha de Origem", font=("Roboto", 16, "bold"), pad=((0,0), (15, 5)), text_color=ACCENT_COLOR)],
        [sg.Text("Defina a planilha de onde os dados serão copiados", font=("Roboto", 10), text_color='#B0B0B0')],
        [sg.HorizontalSeparator(color='#505050')],
        
        [sg.Text("Nome da Planilha (Ex: Relatório Mensal)", font=("Roboto", 10, "bold"), text_color=ACCENT_COLOR, pad=((0,0),(15,5)))],
        [sg.Input(key="-CONFIG_NOME-", size=(55, 1), font=("Consolas", 11), border_width=0)],
        
        [sg.Text("DADOS DA ORIGEM", font=("Roboto", 12, "bold"), text_color='#FFD54F', pad=((0,0),(20,5)))], # Amarelo Suave
        
        [sg.Text("ID da Planilha:", font=("Roboto", 10, "bold"))], 
        [sg.Input(key="-ORIGEM_ID-", size=(55, 1), font=("Consolas", 11), border_width=0)],
        
        [sg.Text("Nome da Aba:", font=("Roboto", 10, "bold"))], 
        [sg.Input(key="-ORIGEM_ABA-", size=(55, 1), font=("Consolas", 11), border_width=0)],
        
        [sg.Text("")],
        [sg.Column([
            [sg.Button("SALVAR", key="-SAVE-", size=(15, 1), button_color=(COLOR_BOTON_START, BTN_SUCCESS), font=("Roboto", 10, "bold"), border_width=0),
             sg.Button("CANCELAR", key="-CANCEL-", size=(15, 1), button_color=(COLOR_BOTON_START, BTN_DANGER), font=("Roboto", 10, "bold"), border_width=0)]
        ], justification='center', pad=(0,15))]
    ]
    
    return sg.Window("Nova Planilha", layout, modal=True, size=(500, 420), keep_on_top=True)

def criar_modal_novo_destino():
    """Cria modal para adicionar novo destino"""
    sg.theme('DarkGrey14')
    
    # Cores
    COLOR_BOTON_START= '#E0E0E0'
    BG_COLOR = '#2D2D2D'
    INPUT_BG = '#383838'
    TEXT_COLOR = '#E0E0E0'
    BTN_SUCCESS = '#4CAF50'
    BTN_DANGER = '#E57373'
    ACCENT_COLOR = '#64B5F6'
    
    sg.set_options(background_color=BG_COLOR, text_element_background_color=BG_COLOR, element_background_color=BG_COLOR, input_elements_background_color=INPUT_BG, input_text_color=TEXT_COLOR, text_color=TEXT_COLOR)
    
    layout = [
        [sg.Text("➕ Planilha de Destino", font=("Roboto", 16, "bold"), pad=((0,0), (15, 5)), text_color=ACCENT_COLOR)],
        [sg.Text("Adicione uma planilha para receber os dados", font=("Roboto", 10), text_color='#B0B0B0')],
        [sg.HorizontalSeparator(color='#505050')],
        
        [sg.Text("Nome da Planilha (Ex: Financeiro)", font=("Roboto", 10, "bold"), text_color=ACCENT_COLOR, pad=((0,0),(15,5)))],
        [sg.Input(key="-DEST_NOME-", size=(55, 1), font=("Consolas", 11), border_width=0)],
        
        [sg.Text("DADOS DO DESTINO", font=("Roboto", 12, "bold"), text_color='#FFD54F', pad=((0,0),(20,5)))],
        
        [sg.Text("ID da Planilha:", font=("Roboto", 10, "bold"))],
        [sg.Input(key="-DEST_ID-", size=(55, 1), font=("Consolas", 11), border_width=0)],
        
        [sg.Text("Nome da Aba:", font=("Roboto", 10, "bold"))],
        [sg.Input(key="-DEST_ABA-", size=(55, 1), font=("Consolas", 11), border_width=0)],
        
        [sg.Text("")],
        [sg.Column([
            [sg.Button("ADICIONAR", key="-SAVE-DEST-", size=(15, 1), button_color=(COLOR_BOTON_START, BTN_SUCCESS), font=("Roboto", 10, "bold"), border_width=0),
             sg.Button("CANCELAR", key="-CANCEL-", size=(15, 1), button_color=(COLOR_BOTON_START, BTN_DANGER), font=("Roboto", 10, "bold"), border_width=0)]
        ], justification='center', pad=(0,15))]
    ]
    
    return sg.Window("Novo Destino", layout, modal=True, size=(500, 420), keep_on_top=True)

def criar_janela_principal(config):
    """Cria janela principal do aplicativo"""
    sg.theme('DarkGrey14')
    
    # Palette Cinza 
    COLOR_BOTON_START= '#E0E0E0'
    COLOR_LOG_SISTEM= '#E0E0E0'
    COLOR_AUTOMACAO = '#E0E0E0'
    COLOR_BG = '#2D2D2D'       # Cinza Fundo
    COLOR_FRAME = '#383838'    # Cinza Frames
    COLOR_TEXT = '#E0E0E0'     # Branco Suave
    COLOR_SUBTEXT = '#B0B0B0'  # Cinza Texto Secundário
    COLOR_PRIMARY = '#4A90E2'  # Azul Corporativo
    COLOR_SECONDARY = '#FFD54F'# Amarelo
    COLOR_SUCCESS = '#4CAF50'  # Verde
    COLOR_DANGER = '#E57373'   # Vermelho
    COLOR_ACCENT = '#64B5F6'   # Azul Claro
    
    # Fontes
    FONT_TITLE = ("Roboto", 22, "bold")
    FONT_SUBTITLE = ("Roboto", 10)
    FONT_HEADER = ("Roboto", 11, "bold")
    FONT_BODY = ("Roboto", 10)
    FONT_MONO = ("Consolas", 10)

    # Configs Globais
    sg.set_options(
        background_color=COLOR_BG, 
        text_element_background_color=COLOR_BG, 
        element_background_color=COLOR_BG,
        input_elements_background_color=COLOR_FRAME,
        input_text_color=COLOR_TEXT,
        text_color=COLOR_TEXT,
        scrollbar_color=COLOR_FRAME,
        button_color=(COLOR_BG, COLOR_PRIMARY)
    )
    
    # Lista de configurações
    planilhas = config.get("planilhas", [])
    nomes_planilhas = [p["nome"] for p in planilhas] if planilhas else ["Nenhuma planilha"]
    
    layout = [
        # --- HEADER ---
        [sg.Column([
            [sg.Text("⚡ PLANILHADOR PRO", font=FONT_TITLE, text_color=COLOR_PRIMARY, pad=(0,0))],
            [sg.Text("Automação de Transferência de Dados", font=FONT_SUBTITLE, text_color=COLOR_AUTOMACAO, pad=(0,0))]
        ], pad=((20,0), (20, 10)))],
        
        # --- SELEÇÃO DE ORIGEM ---
        [sg.Frame("", [
            [sg.Text("SELECIONE A PLANILHA", font=FONT_HEADER, text_color=COLOR_ACCENT)],
            [sg.Combo(nomes_planilhas, key="-SELECT_CONFIG-", size=(45, 1), readonly=True, enable_events=True, 
                      default_value=nomes_planilhas[0] if nomes_planilhas[0] != "Nenhuma configuração" else None, 
                      font=("Roboto", 12), text_color='black', background_color='white'), # Combo nativo windows
             sg.Button("NOVA PLANILHA", key="-NEW_CONFIG-", button_color=('white', COLOR_PRIMARY), size=(18, 1), font=("Roboto", 10, "bold"), border_width=0),
             sg.Button("REMOVER PLANILHA ", key="-DEL_CONFIG-", button_color=('white', COLOR_DANGER), size=(18, 1), font=("Roboto", 10, "bold"), border_width=0)],
             
            [sg.Text("ℹ️ Selecione uma planilha acima para ver os detalhes", key="-INFO_ORIGEM_ID-", font=FONT_MONO, text_color=COLOR_LOG_SISTEM, size=(80, 1), pad=((5,0),(10,0)))],
        ], pad=((20,20), (0, 15)), border_width=0, background_color=COLOR_BG, expand_x=True, relief=sg.RELIEF_FLAT)],
        
        # --- DESTINOS ---
        [sg.Frame("", [
             [
                 sg.Text("PLANILHAS DE DESTINOS", font=FONT_HEADER, text_color=COLOR_SECONDARY),
                 sg.Push(),
                 sg.Button("ADICIONAR DESTINO", key="-ADD_DESTINO-", button_color=(COLOR_BOTON_START, COLOR_SUCCESS), size=(18, 1), font=("Roboto", 9, "bold"), border_width=0),
                 sg.Button("REMOVER DESTINO", key="-DEL_DESTINO-", button_color=('white', COLOR_DANGER), size=(18, 1), font=("Roboto", 9, "bold"), border_width=0)
             ],
             
             [sg.Listbox(values=[], key="-LISTA_DESTINOS-", size=(80, 6), enable_events=True, font=FONT_MONO, 
                         background_color='#181825', text_color=COLOR_TEXT, no_scrollbar=False, highlight_background_color=COLOR_FRAME, highlight_text_color=COLOR_PRIMARY, expand_x=True)],
        ], pad=((20,20), (0, 15)), border_width=0, background_color=COLOR_BG, expand_x=True, relief=sg.RELIEF_FLAT)],
        
        # --- AÇÕES ---
        [sg.Column([
            [sg.Button("INICIAR CÓPIA", key="-START-", size=(25, 1), button_color=(COLOR_BOTON_START, COLOR_SUCCESS), font=("Roboto", 14, "bold"), border_width=0)]
        ], justification='center', pad=((0,0), (10, 10)))],
        
        # --- PROGRESSO ---
        [sg.Text("STATUS DO PROCESSO", font=FONT_HEADER, text_color=COLOR_TEXT, pad=((25,0),(0,5)))],
        [sg.ProgressBar(100, orientation='h', size=(55, 15), key='-PROGRESSBAR-', bar_color=(COLOR_SUCCESS, '#181825'), expand_x=True, pad=((25,25),(0,5)), border_width=0)],
        [sg.Text("Pronto para iniciar", key="-STATUS-", size=(100, 1), font=FONT_BODY, text_color=COLOR_LOG_SISTEM, pad=((100,100),(0,15)), justification='center')],
        
        # --- TABS ---
        [sg.TabGroup([[
            sg.Tab('  LOG DO SISTEMA  ', [
                [sg.Multiline(size=(80, 10), key="-OUTPUT-", autoscroll=True, disabled=True, font=FONT_MONO, 
                              background_color='#11111B', text_color=COLOR_SUCCESS, expand_x=True, expand_y=True, border_width=0)]
            ], background_color=COLOR_BG, element_justification='c', title_color=COLOR_LOG_SISTEM)
        ]], tab_location='topleft', title_color=COLOR_SUBTEXT, tab_background_color=COLOR_FRAME, selected_title_color='#E0E0E0', 
           selected_background_color=COLOR_PRIMARY, border_width=0, expand_x=True, expand_y=True, pad=((20,10),(0,20)))],
        
        [sg.Push(), sg.Button("FECHAR", key="-EXIT-", size=(12, 1), button_color=(COLOR_TEXT, '#45475A'), font=("Roboto", 9, "bold"), border_width=0, pad=((0,25),(0,15)))]
    ]
    
    return sg.Window("Planilhador Pro", layout, size=(900, 700), finalize=True, resizable=True, icon=ICONE_BASE64)

def atualizar_interface_principal(window, config, nome_config):
    """Atualiza a lista de destinos e informações da origem"""
    planilhas = config.get("planilhas", [])
    if not planilhas:
        window["-INFO_ORIGEM_ID-"].update("Nenhuma configuração criada.")
        window["-LISTA_DESTINOS-"].update([])
        return

    planilha_selecionada = next((p for p in planilhas if p["nome"] == nome_config), None)
    
    if planilha_selecionada:
        # Atualiza Info Origem
        origem = planilha_selecionada.get("origem", {})
        window["-INFO_ORIGEM_ID-"].update(f"ID Origem: {origem.get('id', '')} | Aba: {origem.get('aba', '')}")
        
        # Atualiza Lista Destinos
        destinos = planilha_selecionada.get("destinos", [])
        lista_formatada = [f"{d['nome']} | Aba: {d['aba']}" for d in destinos]
        window["-LISTA_DESTINOS-"].update(lista_formatada)
    else:
        window["-INFO_ORIGEM_ID-"].update("Selecione uma configuração.")
        window["-LISTA_DESTINOS-"].update([])

def run_copy_all_task(window, planilha_config):
    destinos = planilha_config.get("destinos", [])
    total_destinos = len(destinos)
    erros = []
    
    try:
        window.write_event_value('-LOG-', f"\n✅ Autenticando com Google Sheets...")
        creds = autenticar()
        service = build('sheets', 'v4', credentials=creds)
        window.write_event_value('-LOG-', f"\n✅ Autenticado com sucesso!\n\n")
        
        # Limpa resultados anteriores
        window.write_event_value('-CLEAR_RESULTS-', '')
        
        origem_id = planilha_config["origem"]["id"]
        origem_aba = planilha_config["origem"]["aba"]
        
        for idx, destino in enumerate(destinos, 1):
            destino_nome = destino["nome"]
            
            # Callback para atualizar a barra de progresso
            def atualizar_progresso(pct_atual):
                # Calcula o progresso global baseado no arquivo atual
                # (idx-1) completos + fração do atual
                progresso_base = ((idx - 1) / total_destinos) * 100
                progresso_total = progresso_base + (pct_atual / total_destinos)
                window.write_event_value('-PROGRESS-', int(progresso_total))
                window["-STATUS-"].update(f"Colando em {destino_nome} | {pct_atual:.2f}%")

            window.write_event_value('-CURRENT-', f"Colando em {destino_nome} ({idx}/{total_destinos})")
            
            try:
                window.write_event_value('-LOG-', f"\n✅ Colando em {destino_nome}...")
                window.write_event_value('-UPDATE_RESULT_STATUS-', {'nome': destino_nome, 'status': 'wait'})
                
                copiar_para_aba_existente(
                    service,
                    origem_id,
                    origem_aba,
                    destino["id"],
                    destino["aba"],
                    callback_progresso=atualizar_progresso
                )
                
                window.write_event_value('-LOG-', f"\n✅ {destino_nome} - Concluído!\n")
                window.write_event_value('-UPDATE_RESULT_STATUS-', {'nome': destino_nome, 'status': 'success'})
                
            except Exception as e:
                error_msg = f"\n❌ {destino_nome} - ERRO: {str(e)}\n"
                window.write_event_value('-LOG-', error_msg)
                window.write_event_value('-UPDATE_RESULT_STATUS-', {'nome': destino_nome, 'status': 'error'})
                erros.append({
                    "nome": destino_nome,
                    "destino": destino,
                    "erro": str(e)
                })
        
        window.write_event_value('-PROGRESS-', 100)
        
        status_final = 'success'
        if erros:
            window.write_event_value('-LOG-', f"\n✅ Finalizado com {len(erros)} erro(s).")
            status_final = 'error'
        else:
            window.write_event_value('-LOG-', f"\n\n🎉🎉🎉🎉 SUCESSO TOTAL! 🎉🎉🎉🎉")
            
        window.write_event_value('-DONE-', {'status': status_final, 'erros': erros, 'config': planilha_config})
            
    except Exception as e:
        window.write_event_value('-LOG-', f"\n❌ ERRO CRÍTICO: {str(e)}")
        window.write_event_value('-DONE-', {'status': 'critical_error', 'erros': [], 'config': planilha_config})

def main():
    config = load_config()
    window = criar_janela_principal(config)
    
    # Init
    if config.get("planilhas"):
        atualizar_interface_principal(window, config, config["planilhas"][0]["nome"])
    
    while True:
        event, values = window.read()
        
        if event in (sg.WINDOW_CLOSED, "-EXIT-"):
            break
            
        # --- SELEÇÃO DE CONFIGURAÇÃO ---
        if event == "-SELECT_CONFIG-":
            nome_config = values["-SELECT_CONFIG-"]
            if nome_config and nome_config != "Nenhuma configuração":
                atualizar_interface_principal(window, config, nome_config)
        
        # --- NOVA CONFIGURAÇÃO (ORIGEM) ---
        if event == "-NEW_CONFIG-":
            modal = criar_modal_config_origem()
            while True:
                ev_m, val_m = modal.read()
                if ev_m == sg.WINDOW_CLOSED:
                    break
                    
                if ev_m == "-CANCEL-":
                    modal.close()
                    break
                    
                if ev_m == "-SAVE-":
                    if not all([val_m["-CONFIG_NOME-"], val_m["-ORIGEM_ID-"], val_m["-ORIGEM_ABA-"]]):
                        sg.popup_error("Preencha todos os campos!", keep_on_top=True)
                        continue
                        
                    nova_config = {
                        "nome": val_m["-CONFIG_NOME-"],
                        "origem": {"id": val_m["-ORIGEM_ID-"], "aba": val_m["-ORIGEM_ABA-"]},
                        "destinos": []
                    }
                    
                    if "planilhas" not in config: config["planilhas"] = []
                    config["planilhas"].append(nova_config)
                    save_config(config)
                    
                    modal.close() # Fecha modal antes do popup
                    
                    sg.popup("✅ Origem Configurada! Agora adicione os destinos.", title="Sucesso", keep_on_top=True)
                    
                    # Atualiza combo e seleciona a nova
                    nomes_planilhas = [p["nome"] for p in config["planilhas"]]
                    window["-SELECT_CONFIG-"].update(values=nomes_planilhas, value=nova_config["nome"])
                    atualizar_interface_principal(window, config, nova_config["nome"])
                    
                    break


        # --- REMOVER CONFIGURAÇÃO (ORIGEM) ---
        if event == "-DEL_CONFIG-":
            nome_config = values["-SELECT_CONFIG-"]
            if not nome_config or nome_config == "Nenhuma configuração":
                sg.popup_error("Selecione uma configuração para remover.")
                continue

            if sg.popup_yes_no(f"Tem certeza que deseja remover a configuração '{nome_config}'?\nIsso excluirá a origem e todos os seus destinos associados.") == "Yes":
                # Filtra removendo a configuração selecionada
                nova_lista_planilhas = [p for p in config["planilhas"] if p["nome"] != nome_config]
                config["planilhas"] = nova_lista_planilhas
                save_config(config)
                
                sg.popup(f"Configuração '{nome_config}' removida com sucesso!")
                
                # Atualiza a lista na interface
                nomes_planilhas = [p["nome"] for p in config["planilhas"]] if config["planilhas"] else ["Nenhuma configuração"]
                
                # Define o novo valor selecionado (o primeiro da lista ou o placeholder)
                novo_valor = nomes_planilhas[0] if nomes_planilhas else None
                
                window["-SELECT_CONFIG-"].update(values=nomes_planilhas, value=novo_valor)
                
                # Atualiza os detalhes na tela se houver algo selecionado, senão limpa
                if novo_valor and novo_valor != "Nenhuma configuração":
                    atualizar_interface_principal(window, config, novo_valor)
                else:
                    # Limpa manualmente se não sobrou nada
                    window["-INFO_ORIGEM_ID-"].update("Nenhuma configuração criada.")
                    window["-LISTA_DESTINOS-"].update([])

        # --- ADICIONAR DESTINO ---
        if event == "-ADD_DESTINO-":
            nome_config = values["-SELECT_CONFIG-"]
            if not nome_config or nome_config == "Nenhuma configuração":
                sg.popup_error("Selecione ou crie uma configuração de origem primeiro.")
                continue
                
            modal = criar_modal_novo_destino()
            while True:
                ev_d, val_d = modal.read()
                if ev_d in (sg.WINDOW_CLOSED, "-CANCEL-"):
                    break
                if ev_d == "-SAVE-DEST-":
                    if not all([val_d["-DEST_NOME-"], val_d["-DEST_ID-"], val_d["-DEST_ABA-"]]):
                        sg.popup_error("Preencha todos os campos do destino!")
                        continue
                    
                    novo_destino = {
                        "nome": val_d["-DEST_NOME-"],
                        "id": val_d["-DEST_ID-"],
                        "aba": val_d["-DEST_ABA-"]
                    }
                    
                    # Encontra config atual e adiciona
                    for p in config["planilhas"]:
                        if p["nome"] == nome_config:
                            if "destinos" not in p: p["destinos"] = []
                            p["destinos"].append(novo_destino)
                            break
                    
                    save_config(config)
                    atualizar_interface_principal(window, config, nome_config)
                    modal.close()
                    break
            modal.close()
            
        # --- REMOVER DESTINO ---
        if event == "-DEL_DESTINO-":
            nome_config = values["-SELECT_CONFIG-"]
            selecao = values["-LISTA_DESTINOS-"]
            if not selecao:
                sg.popup_error("Selecione um destino na lista para remover.")
                continue
            
            # O formato da string na lista é: "Nome Destino | Aba: NomeAba"
            # Vamos pegar tudo antes do primeiro " | " como o nome
            texto_selecionado = selecao[0]
            nome_destino_selecionado = texto_selecionado.split(" | Aba:")[0]
            
            if sg.popup_yes_no(f"Remover destino '{nome_destino_selecionado}'?") == "Yes":
                 item_removido = False
                 for p in config["planilhas"]:
                        if p["nome"] == nome_config:
                            # Filtra removendo o destino com o nome correspondente
                            nova_lista = [d for d in p["destinos"] if d["nome"] != nome_destino_selecionado]
                            
                            if len(nova_lista) < len(p["destinos"]):
                                p["destinos"] = nova_lista
                                item_removido = True
                            break
                 
                 if item_removido:
                     save_config(config)
                     atualizar_interface_principal(window, config, nome_config)
                     sg.popup(f"Destino '{nome_destino_selecionado}' removido com sucesso!")
                 else:
                     sg.popup_error("Não foi possível encontrar o destino para remover.")

        # --- EXECUÇÃO ---
        if event == "-START-":
            nome_config = values["-SELECT_CONFIG-"]
            if not nome_config: continue
            
            # Pega config completa
            planilha_config = next((p for p in config["planilhas"] if p["nome"] == nome_config), None)
            if not planilha_config or not planilha_config.get("destinos"):
                sg.popup_error("Configure pelo menos um destino antes de iniciar.")
                continue
            
            
            # Setup UI
            window["-START-"].update("CARREGANDO", button_color=('#FFFFFF', '#E57373'), disabled=True, disabled_button_color=('#FFFFFF', '#E57373'))
            window["-OUTPUT-"].update("")
            window["-PROGRESSBAR-"].update(0)
            
            # Prepara tabela de resultados
            # [Nome, Status, Detalhes]
            dados_tabela = []
            for d in planilha_config["destinos"]:
                dados_tabela.append([d['nome'], "⏳ Aguardando", ""])
            
            
            threading.Thread(target=run_copy_all_task, args=(window, planilha_config), daemon=True).start()

        # --- EVENTOS THREAD ---
        if event == "-LOG-":
            window["-OUTPUT-"].update(window["-OUTPUT-"].get() + values[event])
            
        if event == "-PROGRESS-":
             window["-PROGRESSBAR-"].update(values[event])
            
        if event == "-UPDATE_RESULT_STATUS-":
            data = values[event]
            nm = data['nome']
            st = data['status']
        

        if event == "-CLEAR_RESULTS-":
             pass # Tabela já iniciada no START

        if event == "-DONE-":
            window["-START-"].update("INICIAR CÓPIA", button_color=('#E0E0E0', '#4CAF50'), disabled=False)
            window["-STATUS-"].update("Concluído.")

    window.close()

if __name__ == "__main__":
    main()
