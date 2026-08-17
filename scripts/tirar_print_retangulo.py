import time
import shutil
from pathlib import Path
from datetime import datetime
from PIL import ImageGrab
import pyautogui

OUTPUT_DIR = Path(r"c:\Users\SESI\Documents\GitHub\automacao-criacao-titulos\outputs")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

ARTIFACTS_DIR = Path(r"C:\Users\SESI\.gemini\antigravity-ide\brain\f478e23f-324f-4810-b279-6a209e233952")

# Coordenadas especificadas pelo usuário:
# p1 = (493, 436), p2 = (1319, 436), p3 = (493, 600), p4 = (1319, 600)
x1, y1 = 493, 436
x2, y2 = 1319, 600
largura = x2 - x1
altura = y2 - y1

print("=" * 65)
print(" CAPTURADOR DE TELA POR REGIÃO - AUTOMATIZADO")
print(f" Coordenadas: p1=({x1}, {y1}), p2=({x2}, {y1}), p3=({x1}, {y2}), p4=({x2}, {y2})")
print(f" Dimensões do Retângulo: {largura}px de largura x {altura}px de altura")
print("=" * 65)

print("\nAguardando 3 segundos... Mude o foco para a janela desejada!")
for i in range(3, 0, -1):
    print(f" >>> Capturando em {i} segundo(s)...", end="\r", flush=True)
    time.sleep(1)

print("\n\n>>> CAPTURANDO IMAGEM AGORA... <<<")

img = None

# Método 1: ImageGrab com bbox
try:
    img = ImageGrab.grab(bbox=(x1, y1, x2, y2))
except Exception as e1:
    print(f"[Aviso] ImageGrab falhou ({e1}). Tentando PyAutoGUI...")

# Método 2: PyAutoGUI screenshot com region
if img is None:
    try:
        img = pyautogui.screenshot(region=(x1, y1, largura, altura))
    except Exception as e2:
        print(f"[Erro] PyAutoGUI também falhou ({e2}).")

if img is not None:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"print_retangulo_{timestamp}.png"
    
    # Salvar na pasta outputs
    caminho_output = OUTPUT_DIR / filename
    caminho_fixo = OUTPUT_DIR / "print_retangulo.png"
    
    img.save(caminho_output)
    img.save(caminho_fixo)
    
    print("\n" + "=" * 65)
    print(f" ✔ PRINT CAPTURADO COM SUCESSO!")
    print(f" Arquivo salvo em: {caminho_output}")
    print(f" Arquivo fixo: {caminho_fixo}")
    
    # Copiar para a pasta de artefatos
    if ARTIFACTS_DIR.exists():
        try:
            shutil.copy(caminho_output, ARTIFACTS_DIR / filename)
            shutil.copy(caminho_output, ARTIFACTS_DIR / "print_retangulo.png")
            print(f" Cópia nos artefatos: {ARTIFACTS_DIR / 'print_retangulo.png'}")
        except Exception as e_art:
            print(f"[Aviso] Não foi possível copiar para artefatos: {e_art}")
    print("=" * 65 + "\n")
else:
    print("\n[ERRO] Não foi possível capturar a tela. Certifique-se de executar no terminal com a área de trabalho ativa.\n")
