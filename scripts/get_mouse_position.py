"""
get_mouse_position.py
=====================
Ferramenta auxiliar para descobrir as coordenadas (x, y) dos campos
no sistema ERP antes de configurar o config.json.

COMO USAR:
1. Execute este script: python get_mouse_position.py
2. Mova o mouse até o campo desejado no sistema ERP
3. Pressione Ctrl+C para capturar a posição atual
4. Anote a posição e atualize o config.json
5. Pressione qualquer tecla para continuar para o próximo campo
"""

import pyautogui
import time
import sys

print("=" * 55)
print("  CAPTURADOR DE COORDENADAS - Automação Financeiro SESI")
print("=" * 55)
print("\nMova o mouse até o campo desejado no ERP e aguarde.")
print("A posição é capturada automaticamente a cada 2 segundos.")
print("\nPressione Ctrl+C para parar.\n")

try:
    while True:
        x, y = pyautogui.position()
        print(f"\r  Posição atual do mouse → X: {x:>5}  |  Y: {y:>5}   ", end="", flush=True)
        time.sleep(0.2)
except KeyboardInterrupt:
    x, y = pyautogui.position()
    print(f"\n\n  ✔  Posição capturada: X = {x}, Y = {y}")
    print(f'\n  Cole no config.json:')
    print(f'  "coordenada_x": {x},')
    print(f'  "coordenada_y": {y}')
    print("\nExecute novamente para capturar o próximo campo.\n")
