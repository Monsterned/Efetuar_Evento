import pandas as pd
import os
import numpy as np
import pyautogui
from datetime import datetime
import pyautogui
import pyperclip
from datetime import datetime, timedelta
from openpyxl import load_workbook

caminho = os.getcwd() 

def click_image(image_path, confidence=0.9):
    current_dir = os.path.dirname(__file__)  # Diretório atual do script
    caminho_imagem = caminho + r'\IMAGENS'
    image_path = os.path.join(current_dir, caminho_imagem, image_path) 
    while True:
        try:
            position = pyautogui.locateOnScreen(image_path, confidence=confidence)
            if position:
                center_x = position.left + position.width // 2
                center_y = position.top + position.height // 2
                pyautogui.click(center_x, center_y)
                print("Imagem foi encontrada na tela.")
                break
        except Exception as e:
            print("Imagem não encontrada na tela. Aguardando...")
        pyautogui.sleep(1)

click_image('situacao_veiculo.png')
pyautogui.sleep(0.5)
pyautogui.press('down')
pyautogui.press('enter')
