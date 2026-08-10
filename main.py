"""
main.py - Script Único de Automação de Criação de Títulos no Senior ERP
========================================================================
Este script centraliza toda a lógica de automação:
1. Leitura do arquivo config/config.json (coordenadas, mapeamento das filiais e delays).
2. Leitura da fila de títulos tratados em data/titulos_tratados.json.
3. Contagem regressiva de 5 segundos.
4. Troca de filial inteligente (apenas quando a filial muda).
5. Confirmação condicional: O clique em Enter / Sim só ocorre quando há notificação de tela aberta.
6. Sequência de cliques nos menus (maleta > finanças > contas_receber > entrada_manutencao > titulos_manutencao).
7. Fail-Safe: Se o mouse for movido para o canto superior esquerdo (X <= 15, Y <= 15), a automação para IMEDIATAMENTE.
"""

import os
import json
import time
import sys
from pathlib import Path
import pyautogui

# Habilitar o recurso nativo de Fail-Safe do PyAutoGUI
pyautogui.FAILSAFE = True

# Diretores base do projeto
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
CONFIG_DIR = BASE_DIR / "config"
SCRIPTS_DIR = BASE_DIR / "scripts"

# Arquivos de dados e configuração
TITULOS_JSON_PATH = DATA_DIR / "titulos_tratados.json"
CONFIG_JSON_PATH = CONFIG_DIR / "config.json"

# Carregamento de Configurações
def carregar_config():
    if not CONFIG_JSON_PATH.exists():
        raise FileNotFoundError(f"Arquivo de configuração não encontrado em: {CONFIG_JSON_PATH}")
    with open(CONFIG_JSON_PATH, "r", encoding="utf-8") as f:
        return json.load(f)

CONFIG_DATA = carregar_config()
MAPA_FILIAIS = CONFIG_DATA.get("filiais", {})
BOTOES = CONFIG_DATA.get("botoes", {})
DELAYS = CONFIG_DATA.get("delays", {
    "home": 0.8,
    "key": 0.15,
    "enter": 1.2,
    "click": 0.8,
    "passo_menu": 1.0,
    "abrir_janela": 2.0,
    "entre_titulos": 1.5
})

# Carregamento de Dados
def carregar_titulos():
    if not TITULOS_JSON_PATH.exists():
        raise FileNotFoundError(f"Arquivo {TITULOS_JSON_PATH} não encontrado. Execute o script processar_planilha.py primeiro.")
    with open(TITULOS_JSON_PATH, "r", encoding="utf-8") as f:
        return json.load(f)

# Funções Auxiliares de Configuração
def obter_info_filial(codigo):
    codigo_str = str(codigo).strip()
    if codigo_str not in MAPA_FILIAIS:
        raise ValueError(f"Código de filial '{codigo}' não encontrado no config.json.")
    return MAPA_FILIAIS[codigo_str]

def obter_coordenadas_botao(nome_botao):
    nome_key = str(nome_botao).lower().strip()
    if nome_key not in BOTOES:
        raise ValueError(f"Botão '{nome_botao}' não encontrado no config.json.")
    return BOTOES[nome_key]

# Verificação contínua do Fail-Safe (canto superior esquerdo)
def verificar_failsafe():
    x, y = pyautogui.position()
    if x <= 15 and y <= 15:
        raise pyautogui.FailSafeException("Fail-Safe ativado: Mouse posicionado no canto superior esquerdo.")

def aguardar_com_failsafe(tempo_segundos):
    """Substitui o time.sleep para verificar o canto superior esquerdo durante a pausa."""
    passo = 0.05
    tempo_passado = 0.0
    while tempo_passado < tempo_segundos:
        verificar_failsafe()
        time.sleep(passo)
        tempo_passado += passo

# Funções de Automação do ERP
def iniciar_contagem_regressiva(segundos=5):
    print(f"\n=======================================================")
    print(f"[Automação] A automação iniciará em {segundos} segundos.")
    print(f"[Automação] Mude o foco para a janela do Senior ERP agora!")
    print(f"[DICA DE EMERGÊNCIA] Mova o mouse para o CANTO SUPERIOR ESQUERDO para PARAR a qualquer momento.")
    print(f"=======================================================")
    for i in range(segundos, 0, -1):
        verificar_failsafe()
        print(f"  >>> Iniciando em {i} segundo(s)...", end="\r", flush=True)
        aguardar_com_failsafe(1.0)
    print("\n\n[Automação] >>> INICIANDO AUTOMAÇÃO AGORA! <<<\n")

def clicar_botao(nome_botao):
    verificar_failsafe()
    info_btn = obter_coordenadas_botao(nome_botao)
    x, y = info_btn["x"], info_btn["y"]
    desc = info_btn.get("descricao", nome_botao)

    print(f"  [Automação] Clicando no botão '{desc}' em (X: {x}, Y: {y})...")
    pyautogui.click(x, y)
    aguardar_com_failsafe(DELAYS.get("click", 0.8))

def abrir_tela_trocar_filial(tem_tela_aberta=False):
    """
    1. Clica no botão de usuário ('user') para abrir as configurações.
    2. Clica no botão 'trocar_filial'.
    3. APENAS se houver tela aberta (tem_tela_aberta=True), confirma a notificação
       de fechamento de telas ("Deseja fechá-las? [Sim]") pressionando Enter.
    """
    print("  [Automação] Passo 1: Clicando no botão de Usuário...")
    clicar_botao("user")
    aguardar_com_failsafe(DELAYS.get("click", 0.8))

    print("  [Automação] Passo 2: Clicando no botão de Trocar Filial...")
    clicar_botao("trocar_filial")
    aguardar_com_failsafe(1.0)

    if tem_tela_aberta:
        print("  [Automação] Notificação de confirmação detectada. Confirmando fechamento de telas (Enter / Sim)...")
        verificar_failsafe()
        pyautogui.press("enter")
        aguardar_com_failsafe(DELAYS.get("abrir_janela", 2.0))
    else:
        print("  [Automação] Nenhuma tela aberta no ERP. Abrindo seleção de filial diretamente...")
        aguardar_com_failsafe(DELAYS.get("abrir_janela", 1.5))

def selecionar_filial(codigo, pressionar_enter=True):
    codigo_str = str(codigo).strip()
    info = obter_info_filial(codigo_str)
    passos = info["posicao"]
    fantasia = info.get("fantasia", "")

    print(f"  [Automação] Passo 3: Selecionando Filial {codigo_str} - {fantasia} (Posição {passos})...")

    verificar_failsafe()
    pyautogui.press("home")
    aguardar_com_failsafe(DELAYS.get("home", 0.8))

    for _ in range(passos):
        verificar_failsafe()
        pyautogui.press("down")
        aguardar_com_failsafe(DELAYS.get("key", 0.15))

    aguardar_com_failsafe(0.3)

    if pressionar_enter:
        verificar_failsafe()
        pyautogui.press("enter")
        aguardar_com_failsafe(DELAYS.get("enter", 1.2))

    print(f"  [Automação] Filial {codigo_str} selecionada com sucesso.")

def navegar_para_tela_titulos():
    print("\n  [Automação] Navegando pelos menus até a tela principal de Títulos...")
    botoes_sequencia = [
        "maleta",
        "financas",
        "contas_receber",
        "entrada_manutencao",
        "titulos_manutencao"
    ]
    
    for idx, btn in enumerate(botoes_sequencia, start=1):
        verificar_failsafe()
        info = obter_coordenadas_botao(btn)
        print(f"  [Automação] Menu ({idx}/5): {info.get('descricao')} [{btn}]")
        clicar_botao(btn)
        aguardar_com_failsafe(DELAYS.get("passo_menu", 1.0))

    aguardar_com_failsafe(DELAYS.get("abrir_janela", 2.0))
    print("  [Automação] Tela principal de Títulos/Manutenção alcançada com sucesso.")

def executar_automacao_completa(modo_teste=True):
    print("=======================================================")
    print("      AUTOMAÇÃO DE CRIAÇÃO DE TÍTULOS - SENIOR ERP     ")
    print("=======================================================")

    titulos = carregar_titulos()
    if not titulos:
        print("[Automação] Nenhum título para processar.")
        return

    print(f"[Automação] Total de títulos a processar: {len(titulos)}")

    try:
        # 1. Contagem regressiva de 5 segundos
        iniciar_contagem_regressiva(5)

        filial_atual = None
        tela_aberta = False  # Indica se existe tela (ex: F301TCR) aberta no ERP

        # 2. Iteração pelos títulos
        for index, titulo in enumerate(titulos, start=1):
            verificar_failsafe()

            num_titulo = titulo.get("numero_titulo")
            cliente = titulo.get("cliente")
            projeto_filial = str(titulo.get("projeto")).strip()
            valor = titulo.get("valor_original")
            vencimento = titulo.get("data_vencimento")
            obs = titulo.get("observacao")
            centro_custo = titulo.get("centro_custo")

            print(f"\n-------------------------------------------------------")
            print(f"[Título {index}/{len(titulos)}] N° Título: {num_titulo}")
            print(f"  • Cliente: {cliente}")
            print(f"  • Filial (Projeto): {projeto_filial}")
            print(f"  • Valor: R$ {valor}")
            print(f"  • Vencimento: {vencimento}")
            print(f"  • Centro de Custo: {centro_custo}")
            print(f"  • Observação: {obs}")

            # Troca de filial se necessário
            if filial_atual != projeto_filial:
                print(f"\n  ➜ [TROCA DE FILIAL] Filial anterior: '{filial_atual}' → Nova Filial: '{projeto_filial}'")
                abrir_tela_trocar_filial(tem_tela_aberta=tela_aberta)
                selecionar_filial(projeto_filial)
                navegar_para_tela_titulos()
                filial_atual = projeto_filial
                tela_aberta = True  # Tela F301TCR passa a estar aberta
            else:
                print(f"\n  ➜ [MANTER FILIAL] Filial atual já é '{filial_atual}'. Clicando em Títulos/Manutenção...")
                clicar_botao("titulos_manutencao")

            if modo_teste:
                print(f"  ✔ [CONSOLE TESTE] Dados do título {num_titulo} exibidos no console.")

            aguardar_com_failsafe(DELAYS.get("entre_titulos", 1.5))

        print("\n=======================================================")
        print(f"[Automação] Processamento concluído com sucesso!")
        print(f"[Automação] Total de {len(titulos)} títulos processados.")
        print("=======================================================")

    except pyautogui.FailSafeException:
        print("\n\n" + "!" * 65)
        print(" INTERRUPÇÃO DE EMERGÊNCIA DISPARADA! ")
        print(" Mouse movido para o canto superior esquerdo (Fail-Safe).")
        print(" A automação foi PARADA IMEDIATAMENTE por segurança.")
        print("!" * 65 + "\n")
    except KeyboardInterrupt:
        print("\n\n[Automação] Interrompido pelo usuário via teclado (Ctrl+C).")

if __name__ == "__main__":
    executar_automacao_completa()
