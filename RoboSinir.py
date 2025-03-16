import threading
from time import sleep
from selenium.common import TimeoutException
from selenium.webdriver.support import expected_conditions as EC
import tkinter as tk
from tkinter import messagebox
import logging
from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
import openpyxl
import pandas as pd
import smtplib
from email.message import EmailMessage
from imbox import Imbox
from selenium.webdriver import ActionChains, Keys
import time
import pyautogui
import os
from dotenv import load_dotenv

# Carregar variáveis do .env
load_dotenv()
SITE = os.getenv("SITE")
CNPJ = os.getenv("CNPJ")
CPF = os.getenv("CPF")
SENHA = os.getenv("SENHA")
EMAIL = os.getenv("EMAIL")
SENHADEAPPEMAIL = os.getenv("SENHADEAPPEMAIL")


# INICIALIZAÇÃO DO WEBDRIVER PARA ABRIR O CHROME.
servico = Service(ChromeDriverManager().install())
navegador = webdriver.Chrome(service=servico)

# ABRE A PÁGINA DO SINIR.
navegador.get(SITE)
sleep(5)
pyautogui.hotkey('enter')
print("ola")

# INSERE AS INFORMÇÕES COMO LOGIN E SENHA DO USUÁRIO.
navegador.find_element(By.XPATH, '//*[@id="mat-input-0"]').send_keys(CNPJ)
navegador.find_element(By.XPATH, '//*[@id="mat-input-1"]').send_keys(CPF)
navegador.find_element(By.XPATH, '//*[@id="mat-input-2"]').send_keys(SENHA)

sleep(5)
navegador.find_element(By.XPATH, '//button[@class="mat-raised-button mat-primary"]').click()

navegador.maximize_window()

navegador.execute_script("document.body.style.zoom='50%'")


hover = WebDriverWait(navegador, 10).until(
                EC.presence_of_element_located((By.XPATH, '/html/body/app-root/app-navegacao/mat-sidenav-container/mat-sidenav-content/p-menubar/div/p-menubarsub/ul/li[2]')))


act = ActionChains(navegador)
act.move_to_element(hover).perform()
sleep(1)

navegador.find_element(By.XPATH,
                       '/html/body/app-root/app-navegacao/mat-sidenav-container/mat-sidenav-content/p-menubar/div/p-menubarsub/ul/li[2]/p-menubarsub/ul/li[4]/a/span').click()
sleep(1.5)
# FAZ A LEITURA DA PLANILHA.
try:
    DataFrame = pd.read_excel('sinir.xlsx', sheet_name='Planilha1')
except Exception as e:
    mensagem_erro = f"Não encontrei a Planilha! ai fica dificíl. {e}"
    logging.error(mensagem_erro)
    messagebox.showerror("Erro", mensagem_erro)

# CRIA UMA PLANILHA PARA SALVAR OS DADOS QUE FOREM ENCONTRADOS NAS CONDIÇÕES ABAIXO

workbook = openpyxl.Workbook()
clientes_separados = workbook.active
clientes_separados['A1'] = 'TICKET'
clientes_separados['B1'] = 'CLIENTE/FORNEC.'
clientes_separados['C1'] = 'MTR'
clientes_separados['D1'] = 'PESO LIQ.'
clientes_separados['E1'] = 'DATA S'
clientes_separados['F1'] = 'MOTIVO'

# CRIANDO A INTERFACE PARA ACOMPANHAR A PLANILHA
df = pd.read_excel("sinir.xlsx")

# Criando a interface
root = tk.Tk()
root.title("Progresso da Automação")
label = tk.Label(root, text="Iniciando...", font=("Arial", 14))
label.pack(pady=20)
root.attributes("-topmost", True)



# FUNÇÃO QUE ENVIA E-MAIL AO TERMINAR AS BAIXAS  ATRAVÉS DO GMAIL
def enviar_email():
    host = 'imap.gmail.com'
    email = EMAIL
    password = SENHADEAPPEMAIL

    with Imbox(host, username=email, password=password) as imbox:
        print('Conexão estabelecida com sucesso!')

    # Criando a mensagem
    msg = EmailMessage()
    msg['Subject'] = 'E-mail automático do Wall-e'
    msg['From'] = email
    msg['To'] = ['DESTINATARIO', 'DESTINATARIO2']
    msg.set_content("""
                            <p> Olá<br><br>
                                Esse é um E-mail automático do Robô Wall-e.<br><br>
                            Passando pra avisar que terminei a baixa das MTR's.<br><br>
                            Segue em anexo as MTR's com divergências.</p><br>

                            <p>Att,</p>
                            <p><strong>Wall-e</strong></p>
                        """, subtype='html')

    # Adicionando o anexo
    with open('clientes_separados.xlsx', 'rb') as f:
        file_data = f.read()
        file_name = f.name

    msg.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=file_name)

    # Enviando o email
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(email, password)
        smtp.send_message(msg)

    print('Email enviado com sucesso!')

# Variável de controle para parar a automação
executando = True

# Função para interromper a automação
def parar_automacao():
    global executando
    executando = False
    label.config(text="Automação interrompida!")
    root.update_idletasks()


# Função para salvar o progresso
def salvar_progresso(linha):
    with open("progresso.txt", "w") as f:
        f.write(str(linha))

# Função para recuperar o progresso
def recuperar_progresso():
    if os.path.exists("progresso.txt"):
        with open("progresso.txt", "r") as f:
            linha = int(f.read().strip())
            return max(0, linha - 5)  # Garante que não volte para uma linha negativa
    return 0  # Se não existir, começa do início

def aguardar_carregamento(navegador,mtr_atual,tempo_max=30):
    try:
        WebDriverWait(navegador, tempo_max).until(
            EC.invisibility_of_element_located((By.XPATH, '/html/body/app-root/app-carregando/div/mat-progress-spinner'))
        )

    except:

        navegador.refresh()
        time.sleep(5)
        try:
            # Aguarda o campo aparecer na tela antes de interagir
            campo_mtr = WebDriverWait(navegador, 10).until(
                EC.presence_of_element_located((By.XPATH, "//input[@formcontrolname='manNumeroDestinador']"))
            )

            # Clica e insere o valor
            campo_mtr.click()
            time.sleep(1)
            campo_mtr.send_keys(mtr_atual)
            time.sleep(1)
            campo_mtr.send_keys(Keys.RETURN)

        except Exception as e:
            print(f" Erro ao encontrar o campo MTR após recarga: {e}")



# Recuperar última linha processada
ultima_linha = recuperar_progresso()

# Botão para parar a automação

btn_parar = tk.Button(root, text="Parar Automação", command=parar_automacao, font=("Arial", 12), bg="red", fg="white")
btn_parar.pack(pady=10)


# Converter a coluna MTR para garantir que apenas valores numéricos sejam usados corretamente
DataFrame["MTR"] = pd.to_numeric(DataFrame["MTR"], errors='coerce')  # Converte para número, substituindo erros por NaN

# Se for NaN, substitui por string vazia; caso contrário, converte para inteiro e depois string
DataFrame["MTR"] = DataFrame["MTR"].apply(lambda x: str(int(x)) if pd.notna(x) else "")

# Criar a coluna CONTADOR contando os caracteres corretamente
DataFrame["CONTADOR"] = DataFrame["MTR"].apply(len)

# PERCORRE A PLANILHA BUSCANDO AS INFORMAÇÕES.

def processar_planilha():

    global executando
    for i in range(ultima_linha, len(DataFrame)):  #COMEÇA DA ULTIMA LINHA QUE FOI SALVA
        if not executando:
            salvar_progresso(i)  # Salva o progresso antes de parar
            break  # Sai do loop se o botão for pressionado

        label.config(text=f"Processando linha {i + 1} de {len(df)}...")
        root.update_idletasks()  # Atualiza a interface em tempo real
        time.sleep(1)


        # Obtém valores da planilha
        mtr = DataFrame.loc[i, "MTR"]
        motorista = DataFrame.loc[i, "MOTORISTA"]
        placa = DataFrame.loc[i, "PLACA"]
        data = DataFrame.loc[i, "DATA S"]
        peso = DataFrame.loc[i, "PESO LIQ."]
        ticket = DataFrame.loc[i, "TICKET"]
        tratamento = DataFrame.loc[i, "TECNOLOGIA"]
        cliente = DataFrame.loc[i, "CLIENTE/FORNEC."]
        complementar = str(DataFrame.loc[i, "MTR's COMPLEMENTARES"])  # Converte para string para evitar erro
        contador = DataFrame.loc[i, "CONTADOR"]


        if (cliente == "PETROLEO BRASILEIRO S A PETROBRAS" or cliente == "NACIOPETRO DISTRIB DE PETROLEO LTD" or
                cliente == "SOLVI ESSENCIS" or cliente == "SOLVI ESSENCIS AMBIENTAL S.A."):
            motivo= "Separado pelo nome do Cliente"
            clientes_separados.append([ticket, cliente, mtr, peso, data, motivo])
            sleep(1)
        elif (contador != 12 or contador != 12.0 or mtr == "nan"):

            motivo = "MTR com o tamanho Errado"
            clientes_separados.append([ticket, cliente, mtr, peso, data, motivo])
        else:


            # Lista para armazenar os MTRs que serão inseridos

            mtrs_list = [str(mtr)]  # Começa com o MTR principal

            # Verifica se há MTR's complementares e filtra os que possuem "41", "42" ou "51"

            if pd.notna(complementar) and complementar.strip() and complementar.lower() != "nan":

                # Divide pelo separador "/" e remove espaços desnecessários

                mtrs_complementares = [m.strip().replace("/", "") for m in complementar.split() if
                                       m.strip().startswith(("41", "42", "51"))]

                # Se encontrou MTR's complementares, adiciona à lista

                if mtrs_complementares:
                    mtrs_list.extend(mtrs_complementares)
                    peso_final = peso / len(mtrs_list)  # Divide o peso igualmente
                else:
                    # Se não houver complementares válidos, usa o peso original

                    peso_final = peso
            else:
                # Se a célula estiver vazia ou for inválida, usa o peso original

                peso_final = peso  # Se a célula estiver vazia ou for inválida, usa o peso original

            # Formata o peso para manter apenas uma casa decimal

            peso_formatado = "{:.1f}".format(peso_final)

            # Insere cada MTR individualmente no site
            for mtr_atual in mtrs_list:
                try:
                    # VERIFICA SE O SITE TRAVOU OU SE CAIU A INTERNET



                    campo_mtr = WebDriverWait(navegador, 10).until(
                        EC.element_to_be_clickable((By.XPATH, "//input[@formcontrolname='manNumeroDestinador']")) #XPATH do campo da MTR

                    )
                    campo_mtr.clear()  # Limpa o campo antes de inserir o novo valor
                    campo_mtr.send_keys(mtr_atual)  # Insere o MTR
                    sleep(0.5)

                    # Pressiona Enter para confirmar

                    campo_mtr.send_keys(Keys.RETURN)

                    # Aguarda o carregamento antes de continuar
                    sleep(0.5)
                    # VERIFICA SE O SITE TRAVOU OU SE CAIU A INTERNET

                    aguardar_carregamento(navegador,mtr_atual)




                    """ESPERA O STATUS FICA DISPONÍVEL E SALVA O MESMO SE ELE NÃO CONSEGUIR SALVAR O STATUS SIGNIFICA QUE NÃO
                     FOI ENCONTRADO NENHUM REGISTRO ENTÃO ELE JÁ SALVA E NA PLANILHA E PULA PARA O PRÓXIMO."""

                    try:
                        status_element = navegador.find_element(By.XPATH,
                                                                '/html/body/app-root/app-navegacao/mat-sidenav-container/mat-sidenav-content/'
                                                                'app-meus-mtrs/mat-sidenav-container/mat-sidenav-content/mat-card[3]/form/p-table/div/'
                                                                'div/table/tbody/tr/td[5]')

                        # PEGAR O STATUS

                        status_text = status_element.text

                        # VALIDAR O STATUS E O CLIENTE
                        if (status_text == "Cancelado"
                                or status_text == "Cancelado Automático por prazo" or status_text == 'Armaz Temporário'
                                or peso > 45000):

                            motivo="Status Cancelado ou peso maior que 45000"
                            clientes_separados.append([ticket, cliente, mtr, peso, data, motivo])

                        elif status_text == "Recebido":
                            pass
                        else:

                            # CLICAR EM RECEBER MTR
                            try:

                                WebDriverWait(navegador, 10).until(EC.element_to_be_clickable(
                                        (By.XPATH, '//a[@class="mat-icon-button mat-primary ng-star-inserted"]'))).click()

                            except Exception as e:
                                logging.error(f"Não foi possível clicar no botão de receber a MTR: {e}")

                            sleep(1)

                            # VERIFICA SE TEM DOIS CAMPOS PARA INSERIR O PESO E O TRATAMENTO.

                            try:
                                lapis1 = navegador.find_element(By.XPATH,
                                                                '/html/body/app-root/app-navegacao/mat-sidenav-container/mat-sidenav-content/app-meus-mtrs/mat-sidenav-container/'
                                                                'mat-sidenav-content/p-dialog[2]/div/div[2]/form/mat-card/p-table/div/div[2]/table/tbody/tr[1]/td[6]/a/i')

                                lapis2 = navegador.find_element(By.XPATH,
                                                                '/html/body/app-root/app-navegacao/mat-sidenav-container/mat-sidenav-content/app-meus-mtrs/'
                                                                'mat-sidenav-container/mat-sidenav-content/p-dialog[2]/div/div[2]/form/mat-card/p-table/div/div[2]/table/tbody/tr[2]/td[6]/a/i')

                                pyautogui.press('esc')
                                sleep(1)

                                # APERTA EM CANCELAR SE A CASO ENCONTRAR DOIS CAMPOS PARA INSERIR O PESO E O TRATAMENTO.

                                navegador.find_element(By.XPATH,
                                                       '/html/body/app-root/app-navegacao/mat-sidenav-container/mat-sidenav-content/app-meus-mtrs/mat-sidenav-container/mat-sidenav-content/'
                                                       'p-dialog[2]/div/div[3]/p-footer/div/div[2]/button/span').click()

                                motivo="Dois campos para inserir o Tratamento"
                                clientes_separados.append([ticket, cliente, mtr, peso, data,motivo])
                            except:

                                # LIMPA O CAMPO MOTORISTA E COLA O NOME DO MOTORISTA

                                navegador.find_element(By.XPATH, '//input[@id="mat-input-4"]').clear()
                                sleep(1)

                                navegador.find_element(By.XPATH, '//input[@id="mat-input-4"]').send_keys(str(motorista))
                                sleep(1)

                                # LIMPA O CAMPO DA E PREENCHE A PLACA

                                navegador.find_element(By.XPATH, '//input[@id="mat-input-5"]').clear()
                                sleep(1)

                                campo_placa = navegador.find_element(By.XPATH, '//input[@id="mat-input-5"]')
                                sleep(1)
                                campo_placa.send_keys(placa)
                                sleep(1)
                                campo_placa.send_keys(Keys.TAB)

                                sleep(1)

                                # LIMPA O CAMPO DA E PREENCHE DATA


                                navegador.find_element(By.XPATH, '//input[@id="mat-input-6"]').send_keys(data)
                                sleep(1)

                                # SELECIONAR RESPONSÁVEL

                                navegador.find_element(By.XPATH, '//div[@class="col-g-4"]').click()
                                sleep(1)

                                # CLICAR NO CHECK DO RESPONSÁVEL

                                navegador.find_element(By.XPATH, '//a[@title="Selecionar"]').click()
                                sleep(1)

                                # ESCOLHER O TRATAMENTO
                                actions = ActionChains(navegador)


                                try:

                                    if tratamento == '2402-COMPOSTAGEM DE RES PRIVADOS':  # COMPOSTAGEM
                                        navegador.find_element(By.XPATH, '//div[@class="mat-select-arrow"]').click()
                                        sleep(1)
                                        pyautogui.press('up', presses=36)
                                        sleep(1)
                                        pyautogui.press('down', presses=11)
                                        sleep(1)
                                        actions.send_keys(Keys.ENTER).perform()
                                        sleep(1)

                                    elif tratamento == '2303-AT CLASSE II - RES PRIVADOS':  # Aterro Residuos Classes IIA e IIB

                                        navegador.find_element(By.XPATH, '//div[@class="mat-select-arrow"]').click()
                                        sleep(1)
                                        pyautogui.press('up', presses=36)
                                        sleep(1)
                                        pyautogui.press('down', presses=2)
                                        sleep(1)
                                        actions.send_keys(Keys.ENTER).perform()
                                        sleep(1)
                                    elif tratamento == 'AT CLASSE II - RES PRIVADOS':  # Aterro Residuos Classes IIA e IIB

                                        navegador.find_element(By.XPATH, '//div[@class="mat-select-arrow"]').click()
                                        sleep(1)
                                        pyautogui.press('up', presses=36)
                                        sleep(1)
                                        pyautogui.press('down', presses=2)
                                        sleep(1)
                                        actions.send_keys(Keys.ENTER).perform()
                                        sleep(1)

                                    elif tratamento == 'Manufatura Reversa':  # OUTROS

                                        navegador.find_element(By.XPATH, '//div[@class="mat-select-arrow"]').click()
                                        sleep(1)
                                        pyautogui.press('up', presses=36)

                                        sleep(1)
                                        pyautogui.press('down', presses=23)
                                        sleep(1)
                                        actions.send_keys(Keys.ENTER).perform()
                                        sleep(1)

                                    elif tratamento == '2401-COPROCESSAMENTO RES PRIVADOS':  # BLENDAGEM PARA COPROCESSAMENTO

                                        navegador.find_element(By.XPATH, '//div[@class="mat-select-arrow"]').click()
                                        sleep(1)
                                        pyautogui.press('up', presses=36)
                                        sleep(1)
                                        pyautogui.press('down', presses=9)
                                        sleep(1)
                                        actions.send_keys(Keys.ENTER).perform()
                                        sleep(1)
                                    elif tratamento == 'COPROCESSAMENTO RES PRIVADOS':  # BLENDAGEM PARA COPROCESSAMENTO

                                        navegador.find_element(By.XPATH, '//div[@class="mat-select-arrow"]').click()
                                        sleep(1)
                                        pyautogui.press('up', presses=36)
                                        sleep(1)
                                        pyautogui.press('down', presses=9)
                                        sleep(1)
                                        actions.send_keys(Keys.ENTER).perform()
                                        sleep(1)

                                    elif tratamento == 'TRANSP EFLUENTES PRIVADOS':  # TRATAMENTO DE EFLUENTE

                                        navegador.find_element(By.XPATH, '//div[@class="mat-select-arrow"]').click()
                                        sleep(1)
                                        pyautogui.press('up', presses=36)
                                        sleep(1)
                                        pyautogui.press('down', presses=29)
                                        sleep(1)
                                        actions.send_keys(Keys.ENTER).perform()
                                        sleep(1)

                                    elif tratamento == '2302-AT CLASSE I - RES PRIVADOS':  # Aterro Resíduos Classe I

                                        navegador.find_element(By.XPATH, '//div[@class="mat-select-arrow"]').click()
                                        sleep(1)
                                        pyautogui.press('up', presses=36)
                                        sleep(1)
                                        pyautogui.press('down', presses=1)

                                        sleep(1)
                                        actions.send_keys(Keys.ENTER).perform()
                                        sleep(1)

                                    elif tratamento == '2406 - SERVICOS DE LOGISTICA REVERSA':  # LÂMPADAS

                                        motivo="Lâmpadas"
                                        clientes_separados.append([ticket, cliente, mtr, peso, data,motivo])

                                    # Insere o peso correspondente
                                    campo_peso = WebDriverWait(navegador, 10).until(
                                        EC.element_to_be_clickable((By.XPATH, '//input[@formcontrolname="marQuantidadeRecebida"]'))
                                    )
                                    campo_peso.clear()

                                    # Se for o MTR principal e não houver complementares, mantém o peso original
                                    if mtr_atual == str(mtr) and len(mtrs_list) == 1:
                                        campo_peso.send_keys("{:.1f}".format(peso))
                                    else:
                                        campo_peso.send_keys(peso_formatado)

                                        sleep(1)

                                    # CLICAR NO LAPIS DA JUSTIFICATIVA

                                    navegador.find_element(By.XPATH,
                                                           '/html/body/app-root/app-navegacao/mat-sidenav-container/mat-sidenav-content/app-meus-mtrs/mat-sidenav-container/mat-sidenav-content/p-dialog[2]/div/div[2]/form/mat-card/p-table/div/div[2]/table/tbody/tr/td[6]/a/i').click()
                                    sleep(1)

                                    # PREENCHER O TICKET NA JUSTIFICATIVA

                                    navegador.find_element(By.XPATH, '//textarea[@id="mat-input-38"]').send_keys(str(ticket))
                                    sleep(1)

                                    # CLICAR EM SALVAR O TICKET

                                    navegador.find_element(By.XPATH,
                                                           '//button[@label="Salvar"]').click()
                                    sleep(1)

                                    #  ANTES DE CLICAR EM RECEBER VERIFICA SE O BOTAO DE RECEBER ESTA DISPONIVEL

                                    button = navegador.find_element(By.XPATH, '//button[span[text()="Receber"]]')
                                    is_disabled = button.get_attribute("disabled") is not None

                                    if is_disabled:
                                        # SE CAIU AQUI É PORQUE O BOTÃO NÃO ESTÁ DISPONÍVEL E NÃO PODE SER CLICADO ENTÃO SALVA EM UMA PLANILHA

                                        clientes_separados.append([ticket, cliente, mtr, peso, data])

                                        # print("O botão está desativado e não pode ser clicado.")
                                        # CLICA EM CANCELAER ANTES DE COMEÇAR DE NOVO

                                        navegador.find_element(By.XPATH, '//button[span[text()="Cancelar"]]').click()
                                    else:

                                        # SE CAIU AQUI É PORQUE O BOTÃO ESTÁ DISPONÍVEL ENTÃO CLICA EM RECEBER.
                                        navegador.find_element(By.XPATH,
                                                               '//button[span[text()="Receber"]]').click()

                                    # LIMPAR O CAMPO DA MTR PARA INSERIR A NOVA MTR.
                                    sleep(1)
                                    navegador.find_element(By.XPATH, '//input[@id="mat-input-32"]').click()
                                    sleep(1)
                                    pyautogui.hotkey('ctrl', 'a')
                                    sleep(1)
                                    pyautogui.hotkey('backspace')
                                    sleep(1)

                                except:
                                    navegador.find_element(By.XPATH,
                                                           '/html/body/app-root/app-navegacao/mat-sidenav-container/mat-sidenav-content/app-meus-mtrs/'
                                                           'mat-sidenav-container/mat-sidenav-content/p-dialog[2]/div/div[3]/p-footer/div/div[2]/button/span').click()
                                    motivo="Não conseguiu inserir o Tratamento"

                                    clientes_separados.append([ticket, cliente, mtr, peso, data,motivo])


                    except:
                        motivo="Não conseguiu pegar o Status, provavelmente estava vazio"
                        clientes_separados.append([ticket, cliente, mtr, peso, data,motivo])

                except TimeoutException:
                    print(f"Erro ao inserir MTR {mtr_atual}. POR QUE CAIU AQUI ?...")
                    navegador.refresh()  # Recarrega a página
                    sleep(5)  # Espera antes de tentar novamente

    workbook.save('clientes_separados.xlsx')

    # CHAMA A FUNÇÃO PARA ENVIAR O E-MAIL
    enviar_email()
    
    # Mensagem final quando terminar
    label.config(text="Processo concluído!")


    navegador.quit()  # Fecha o navegador ao final

    label.config(text="Processo concluído!")


# Rodar a automação em uma thread separada para não travar a interface

threading.Thread(target=processar_planilha, daemon=True).start()

root.mainloop()

