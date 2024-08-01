import pandas as pd
import os
import numpy as np
import pyautogui
from datetime import datetime
import pyautogui
import pyperclip
from datetime import datetime, timedelta
from openpyxl import load_workbook


def aviso_atencao(image_path, image_path2,image_path3,image_path4, confidence=0.9):
    current_dir = os.path.dirname(__file__)  # Diretório atual do script
    caminho_imagem = 'IMAGENS'
    image_path = os.path.join(current_dir, caminho_imagem, image_path)
    image_path2 = os.path.join(current_dir, caminho_imagem, image_path2)
    image_path3 = os.path.join(current_dir, caminho_imagem, image_path3)
    image_path4 = os.path.join(current_dir, caminho_imagem, image_path4)
    attempts = 3
    aviso_ativo = False  # Variável para indicar se o aviso está ativo
    aviso_veiculo = False
    # Tentar encontrar a primeira imagem até 3 vezes
    for attempt in range(attempts):
        try:
            position = pyautogui.locateOnScreen(image_path, confidence=confidence)
            if position:
                aviso_ativo = True  # Aviso encontrado, definir como ativo
                break
            else:
                print(f"Imagem 1 não encontrada na tentativa {attempt + 1}. Aguardando...")
        except Exception as e:
            print(f"Erro ao tentar encontrar a imagem 1 na tentativa {attempt + 1}: {e}")

        return aviso_ativo, aviso_veiculo  # Retorna o estado do aviso

aviso_ativo, aviso_veiculo = aviso_atencao('aviso_cnh_vencida.png','ok_marcado.png','ok_efetuado.png','curso_vencido.png')
