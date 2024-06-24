import pandas as pd
import os
import numpy as np
import pyautogui
from datetime import datetime
import pyautogui
import pyperclip
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from datetime import datetime, timedelta
from unidecode import unidecode

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


def alerta_revisonais(image_path,image_path2, confidence=0.9):
    current_dir = os.path.dirname(__file__)  # Diretório atual do script
    caminho_imagem = caminho + r'\IMAGENS'
    image_path = os.path.join(current_dir, caminho_imagem, image_path) 
    image_path2 = os.path.join(current_dir, caminho_imagem, image_path2)
    try:
        position = pyautogui.locateOnScreen(image_path, confidence=confidence)
        if position:
            print("Imagem foi encontrada na tela.")
            click_image('ok_marcado.png')
            caminho_do_arquivo = 'EVENTO.xlsx'
            nome_da_aba = 'Planilha1'
            wb = load_workbook(caminho_do_arquivo)
            ws = wb[nome_da_aba]
            coluna_ost = 'J'  
            if linha > ws.max_row:
                ws[coluna_ost + str(linha_especifica)] = 'VEICULO COM ALERTAS DE REVISAO'
            else:
                ws[coluna_ost + str(linha_especifica)] = 'VEICULO COM ALERTAS DE REVISAO'
            wb.save(caminho_do_arquivo)
            wb.close()
            return True  # Indicar para pular a iteração
        else:
            print("Imagem não encontrada na tela.")
            return False
    except Exception as e:
        print(f"Erro ao tentar encontrar a imagem: {e}")
    try:
        position2 = pyautogui.locateOnScreen(image_path2, confidence=confidence)
        if position2:
            print("Imagem foi encontrada na tela.")
            click_image('ok_marcado.png')
            caminho_do_arquivo = 'EVENTO.xlsx'
            nome_da_aba = 'Planilha1'
            wb = load_workbook(caminho_do_arquivo)
            ws = wb[nome_da_aba]
            coluna_ost = 'J'  
            if linha > ws.max_row:
                ws[coluna_ost + str(linha_especifica)] = 'VEICULO COM ALERTAS DE REVISAO'
            else:
                ws[coluna_ost + str(linha_especifica)] = 'VEICULO COM ALERTAS DE REVISAO'
            wb.save(caminho_do_arquivo)
            wb.close()
            return True  # Indicar para pular a iteração
        else:
            print("Imagem não encontrada na tela.")
            return False
    except Exception as e:
        print(f"Erro ao tentar encontrar a imagem: {e}")

    return False
    pyautogui.sleep(1)


def motorista_nao_localizado(image_path, confidence=0.9):
    current_dir = os.path.dirname(__file__)  # Diretório atual do script
    caminho_imagem = caminho + r'\IMAGENS'
    image_path = os.path.join(current_dir, caminho_imagem, image_path) 
    try:
        position = pyautogui.locateOnScreen(image_path, confidence=confidence)
        if position:
            print("Imagem foi encontrada na tela.")
            caminho_do_arquivo = 'EVENTO.xlsx'
            nome_da_aba = 'Planilha1'
            wb = load_workbook(caminho_do_arquivo)
            ws = wb[nome_da_aba]
            coluna_ost = 'J'  
            if linha > ws.max_row:
                ws[coluna_ost + str(linha_especifica)] = 'MOTORISTA NAO ENCONTRADO'
            else:
                ws[coluna_ost + str(linha_especifica)] = 'MOTORISTA NAO ENCONTRADO'
            wb.save(caminho_do_arquivo)
            wb.close()
            return True  # Indicar para pular a iteração
        else:
            print("Imagem não encontrada na tela.")
            return False
    except Exception as e:
        print(f"Erro ao tentar encontrar a imagem: {e}")
        return False
    pyautogui.sleep(1)


def veiculo_nao_localizado(image_path, confidence=0.9):
    current_dir = os.path.dirname(__file__)  # Diretório atual do script
    caminho_imagem = caminho + r'\IMAGENS'
    image_path = os.path.join(current_dir, caminho_imagem, image_path) 
    try:
        position = pyautogui.locateOnScreen(image_path, confidence=confidence)
        if position:
            print("Imagem foi encontrada na tela.")
            caminho_do_arquivo = 'EVENTO.xlsx'
            nome_da_aba = 'Planilha1'
            wb = load_workbook(caminho_do_arquivo)
            ws = wb[nome_da_aba]
            coluna_ost = 'J'  
            if linha > ws.max_row:
                ws[coluna_ost + str(linha_especifica)] = 'VEICULO NAO ENCONTRADO OU INATIVO'
            else:
                ws[coluna_ost + str(linha_especifica)] = 'VEICULO NAO ENCONTRADO OU INATIVO'
            wb.save(caminho_do_arquivo)
            wb.close()
            return True  # Indicar para pular a iteração
        else:
            print("Imagem não encontrada na tela.")
            return False
    except Exception as e:
        print(f"Erro ao tentar encontrar a imagem: {e}")
        return False
    pyautogui.sleep(1)


def erro_rateio(image_path, confidence=0.9):
    current_dir = os.path.dirname(__file__)  # Diretório atual do script
    caminho_imagem = caminho + r'\IMAGENS'
    image_path = os.path.join(current_dir, caminho_imagem, image_path) 
    try:
        position = pyautogui.locateOnScreen(image_path, confidence=confidence)
        if position:
            print("Imagem foi encontrada na tela.")
            click_image('ok_efetuado.png')
            caminho_do_arquivo = 'EVENTO.xlsx'
            nome_da_aba = 'Planilha1'
            wb = load_workbook(caminho_do_arquivo)
            ws = wb[nome_da_aba]
            coluna_ost = 'J'  
            if linha > ws.max_row:
                ws[coluna_ost + str(linha_especifica)] = 'ERRO DE RATEIO DO EVENTO'
            else:
                ws[coluna_ost + str(linha_especifica)] = 'ERRO DE RATEIO DO EVENTO'
            wb.save(caminho_do_arquivo)
            wb.close()
            return True  # Indicar para pular a iteração
        else:
            print("Imagem não encontrada na tela.")
            return False
    except Exception as e:
        print(f"Erro ao tentar encontrar a imagem: {e}")
        return False
    pyautogui.sleep(1)

def aviso_atencao(image_path, image_path2,image_path3, confidence=0.9):
    current_dir = os.path.dirname(__file__)  # Diretório atual do script
    caminho_imagem = 'IMAGENS'
    image_path = os.path.join(current_dir, caminho_imagem, image_path)
    image_path2 = os.path.join(current_dir, caminho_imagem, image_path2)
    image_path3 = os.path.join(current_dir, caminho_imagem, image_path3)
    attempts = 3
    aviso_ativo = False  # Variável para indicar se o aviso está ativo

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

        pyautogui.sleep(1)

    if not aviso_ativo:
        print("Imagem 1 não encontrada após 3 tentativas. Prosseguindo...")
        return aviso_ativo  # Retorna o estado do aviso
    
    caminho_do_arquivo = 'EVENTO.xlsx'
    nome_da_aba = 'Planilha1'
    wb = load_workbook(caminho_do_arquivo)
    ws = wb[nome_da_aba]
    coluna_ost = 'J'  
    if linha > ws.max_row:
        ws[coluna_ost + str(linha_especifica)] = 'CNH VENCIDA OU COM CARENCIA'
    else:
        ws[coluna_ost + str(linha_especifica)] = 'CNH VENCIDA OU COM CARENCIA'
    wb.save(caminho_do_arquivo)
    wb.close()

    while True:
        try:
            position2 = pyautogui.locateOnScreen(image_path2, confidence=confidence)
            if position2:
                center_x = position2.left + position2.width // 2
                center_y = position2.top + position2.height // 2
                pyautogui.click(center_x, center_y)
                print("Imagem 2 foi encontrada na tela.")
                break
            else:
                print("Imagem 2 não encontrada na tela. Aguardando...")
        except Exception as e:
            print(f"Erro ao tentar encontrar a imagem 2: {e}")

        pyautogui.sleep(1)

    pyautogui.sleep(1)
    while True:
        try:
            position3 = pyautogui.locateOnScreen(image_path3, confidence=confidence)
            if position2:
                center_x = position3.left + position3.width // 2
                center_y = position3.top + position3.height // 2
                pyautogui.click(center_x, center_y)
                print("Imagem 2 foi encontrada na tela.")
                break
            else:
                print("Imagem 2 não encontrada na tela. Aguardando...")
        except Exception as e:
            print(f"Erro ao tentar encontrar a imagem 2: {e}")

        pyautogui.sleep(1)
    
    pyautogui.sleep(1)
    while True:
        try:
            position3 = pyautogui.locateOnScreen(image_path3, confidence=confidence)
            if position2:
                center_x = position3.left + position3.width // 2
                center_y = position3.top + position3.height // 2
                pyautogui.click(center_x, center_y)
                print("Imagem 2 foi encontrada na tela.")
                break
            else:
                print("Imagem 2 não encontrada na tela. Aguardando...")
        except Exception as e:
            print(f"Erro ao tentar encontrar a imagem 2: {e}")

        pyautogui.sleep(1)

    return aviso_ativo  # Retorna o estado do aviso


def aviso_antt(image_path,confidence=0.9):
    current_dir = os.path.dirname(__file__)  # Diretório atual do script
    caminho_imagem = 'IMAGENS'
    image_path = os.path.join(current_dir, caminho_imagem, image_path)

    aviso_antt = False  # Variável para indicar se o aviso está ativo

    try:
        position = pyautogui.locateOnScreen(image_path, confidence=confidence)
        if position:
            aviso_antt = True  # Aviso encontrado, definir como ativo
        else:
            print(f"Imagem não encontrada na tentativa. Aguardando...")
    except Exception as e:
        print(f"Erro ao tentar encontrar a imagem: {e}")

    pyautogui.sleep(1)

    if not aviso_antt:
        print("Imagem  não encontrada. Prosseguindo...")
        return aviso_antt  # Retorna o estado do aviso
    
    caminho_do_arquivo = 'EVENTO.xlsx'
    nome_da_aba = 'Planilha1'
    wb = load_workbook(caminho_do_arquivo)
    ws = wb[nome_da_aba]
    coluna_ost = 'J'  
    if linha > ws.max_row:
        ws[coluna_ost + str(linha_especifica)] = 'ANTT DO VEICULO ESTA VENCENDO'
    else:
        ws[coluna_ost + str(linha_especifica)] = 'ANTT DO VEICULO ESTA VENCENDO'
    wb.save(caminho_do_arquivo)
    wb.close()

    click_image('ok_marcado.png')
    pyautogui.sleep(2)
    return aviso_antt  # Retorna o estado do aviso

def erro_documento_naoencontrato(image_path, image_path2, confidence=0.9):
    current_dir = os.path.dirname(__file__)  # Diretório atual do script
    caminho_imagem = 'IMAGENS'
    image_path = os.path.join(current_dir, caminho_imagem, image_path)
    image_path2 = os.path.join(current_dir, caminho_imagem, image_path2)

    attempts = 3
    documento_nao_localizado = False  # Variável para indicar se o aviso está ativo

    # Tentar encontrar a primeira imagem até 3 vezes
    for attempt in range(attempts):
        try:
            position = pyautogui.locateOnScreen(image_path, confidence=confidence)
            if position:
                documento_nao_localizado = True  # Aviso encontrado, definir como ativo
                break
            else:
                print(f"Imagem 1 não encontrada na tentativa {attempt + 1}. Aguardando...")
        except Exception as e:
            print(f"Erro ao tentar encontrar a imagem 1 na tentativa {attempt + 1}: {e}")

        pyautogui.sleep(1)

    if not documento_nao_localizado:
        print("Imagem 1 não encontrada após 3 tentativas. Prosseguindo...")
        return documento_nao_localizado  # Retorna o estado do aviso

    caminho_do_arquivo = 'EVENTO.xlsx'
    nome_da_aba = 'Planilha1'
    wb = load_workbook(caminho_do_arquivo)
    ws = wb[nome_da_aba]
    coluna_ost = 'J'  
    if linha > ws.max_row:
        ws[coluna_ost + str(linha_especifica)] = 'DOCUMENTO NAO ENCONTRADO'
    else:
        ws[coluna_ost + str(linha_especifica)] = 'DOCUMENTO NAO ENCONTRADO'
    wb.save(caminho_do_arquivo)
    wb.close()
    # Se a primeira imagem foi encontrada, tentar encontrar a segunda imagem
    while True:
        try:
            position2 = pyautogui.locateOnScreen(image_path2, confidence=confidence)
            if position2:
                center_x = position2.left + position2.width // 2
                center_y = position2.top + position2.height // 2
                pyautogui.click(center_x, center_y)
                print("Imagem 2 foi encontrada na tela.")
                break
            else:
                print("Imagem 2 não encontrada na tela. Aguardando...")
        except Exception as e:
            print(f"Erro ao tentar encontrar a imagem 2: {e}")
        pyautogui.sleep(1)

    pyautogui.sleep(0.5)
    click_image('cancelar_inclusao.png')
    return documento_nao_localizado  # Retorna o estado do aviso



def click_info_manifesto(image_path, confidence=0.9):
    current_dir = os.path.dirname(__file__)  # Diretório atual do script
    caminho_imagem = caminho + r'\IMAGENS'
    image_path = os.path.join(current_dir, caminho_imagem, image_path) 
    while True:
        try:
            position = pyautogui.locateOnScreen(image_path, confidence=confidence)
            if position:
                center_x = position.left + position.width // 2
                center_y = position.top + position.height // 2
                pyautogui.moveTo(center_x, center_y)  # Movendo o cursor para a posição da imagem               
                pyautogui.moveRel(40, 0)  # Movendo o cursor para cima
                pyautogui.click()  # Clicando no local da imagem
                print("Imagem foi encontrada na tela.")
                break
        except Exception as e:
            print("Imagem não encontrada na tela. Aguardando...")
        pyautogui.sleep(1)

agora = datetime.now()
um_minuto_atras = agora - timedelta(minutes=1)
agora_formatado = agora.strftime("%d/%m/%Y %H:%M")
um_minuto_atras_formatado = um_minuto_atras.strftime("%H:%M")
#print(um_minuto_atras_formatado)

Planilha_eventos = pd.read_excel("EVENTO.xlsx")
Planilha_eventos['DATA'] = pd.to_datetime(Planilha_eventos['DATA'])
Planilha_eventos['DATA'] = Planilha_eventos['DATA'].dt.strftime('%d/%m/%Y')
# print(Planilha_eventos['DATA'])

linha_especifica = 1

pyautogui.sleep(2)

for i, linha in enumerate(Planilha_eventos.index):
    data = Planilha_eventos.loc[linha, "DATA"]
    filial = Planilha_eventos.loc[linha, "FILIAL"]
    serie = Planilha_eventos.loc[linha, "SERIE"]
    manifesto = Planilha_eventos.loc[linha, "MANIFESTO"]
    codigo_evento = Planilha_eventos.loc[linha, "COD"]
    valor = Planilha_eventos.loc[linha, "VALOR"]
    valor = str(valor)
    valor = valor.replace('.', ',')
    descricao = Planilha_eventos.loc[linha, "DESCRICAO"]
    descricao = str(descricao)
    if 'Ç' in descricao:
        descricao = descricao.replace("Ç", "C")
    elif('ç' in descricao):
        descricao = descricao.replace("ç", "c")
    descricao = unidecode(descricao)
    descricao = descricao.upper()
    placa = Planilha_eventos.loc[linha, "PLACA"]
    linha_especifica += 1

    click_image('incluir.png')
    pyautogui.sleep(2)
    click_info_manifesto('filial.png')
    for i in range(5):
        pyautogui.press('backspace')
    pyautogui.write(str(filial))
    pyautogui.press('tab')
    click_info_manifesto('data_evento.png')
    for i in range(5):
        pyautogui.press('backspace')
    for i in range(15):
        pyautogui.press('del')
    pyautogui.write(str(data))
    pyautogui.write(str(um_minuto_atras_formatado))
    pyautogui.press('tab')
    click_info_manifesto('codigo_evento.png')
    pyautogui.write(str(codigo_evento))
    pyautogui.press('tab')
    click_info_manifesto('campo_veiculo.png')   
    pyautogui.press('F2')
    click_info_manifesto('campo_placa.png')
    pyautogui.write(str(placa))
    click_image('situacao_veiculo.png')
    pyautogui.sleep(0.5)
    for i in range(2):
        pyautogui.press('down')
    pyautogui.press('enter')
    click_image('atualizar.png')
    pyautogui.sleep(2)
    click_image('selecionar.png')
    pyautogui.sleep(2)
    if veiculo_nao_localizado('campo_veiculo_vazio.png'):
        continue  
    for i in range(2):
        pyautogui.press('tab')
    pyautogui.sleep(1)
    #se placa ou motorista der aviso colocar um ok 

    aviso_ativo = aviso_atencao('aviso_cnh_vencida.png','ok_marcado.png','ok_efetuado.png')
    alerta_revisonal = alerta_revisonais('ALERTA_REVISONAIS.png','aviso_dual_avencer.png')
    antt_vencida = aviso_antt('erro_antt.png')
    pyautogui.sleep(1)
    if motorista_nao_localizado('campo_motorista_vazio.png'):
        continue  
    click_info_manifesto('campo_quantidade.png')
    pyautogui.write('1')
    pyautogui.press('tab')
    click_info_manifesto('campo_valor.png')
    pyautogui.write(valor)
    pyautogui.press('tab')
    click_info_manifesto('campo_observacao.png')
    pyautogui.sleep(3)
    pyautogui.write(descricao)
    pyautogui.sleep(2)
    click_image('salvar.png')
    click_info_manifesto('campo_evento.png')
    pyautogui.click(button='right')
    for i in range(3):
        pyautogui.press("down")
    pyautogui.sleep(1)
    pyautogui.press("enter")
    pyautogui.sleep(0.5)
    try:
        text = pyperclip.paste()
        ost = int(text)
        print("Número do EVENTO:", ost)
    except ValueError:
        print("O conteúdo copiado não é um número válido.")
    except Exception as e:
        print("Ocorreu um erro:", str(e))

    caminho_do_arquivo = 'EVENTO.xlsx'
    nome_da_aba = 'Planilha1'
    wb = load_workbook(caminho_do_arquivo)
    ws = wb[nome_da_aba]
    coluna_ost = 'A'  
    if linha > ws.max_row:
        ws[coluna_ost + str(linha_especifica)] = ost
    else:
        ws[coluna_ost + str(linha_especifica)] = ost
    wb.save(caminho_do_arquivo)
    wb.close()

    click_image('inserir_manifesto.png')
    click_image('pasta_amarela.png')
    click_image('selecionar_tipo_documento.png')
    click_image('selecionar_manifesto.png')
    click_image('inserir_filial_manifesto.png')
    pyautogui.write(str(filial))
    pyautogui.sleep(0.5)
    pyautogui.press('tab')
    pyautogui.write(str(serie))
    pyautogui.sleep(0.5)
    pyautogui.press('tab')
    pyautogui.write(str(manifesto))
    pyautogui.sleep(0.5)
    pyautogui.press('tab')
    click_image('setinha_verde.png')
    pyautogui.sleep(2)
    #Verificar se o documento foi encontrado
    documento_nao_localizado = erro_documento_naoencontrato('erro_documento_naoencontrado.png','ok_efetuado.png')
    click_image('botao_voltar.png')
    if documento_nao_localizado:
        print("Documento nao encontrado, ir para o proximo evento.")
    else:
        print("Documento encontrado. Prosseguir com outra ação.")
        if antt_vencida:
            click_image('ok_efetuado.png')
        if alerta_revisonal:
            click_image('ok_efetuado.png')
        if aviso_ativo:
            print("O aviso está ativo. Tome uma ação específica.")
            click_image('ok_efetuado.png')
            pyautogui.sleep(1)
            click_image('ok_efetuado.png')
        else:
            print("O aviso não está ativo. Prosseguir com outra ação.")
        pyautogui.sleep(2)
        click_image('efetuar.png')
        pyautogui.sleep(2)
        if erro_rateio('erro_rateio.png'):
            continue  
        click_image('yes_efetuar.png')
        pyautogui.sleep(1)
        click_image('ok_efetuado.png')
        pyautogui.sleep(1)

        
