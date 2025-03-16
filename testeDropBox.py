
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pyautogui
from time import sleep

# Iniciar o navegador
navegador = webdriver.Chrome()

# Abrir a página (substitua pela URL real)
navegador.get("https://mtr.sinir.gov.br/")

# INSERE AS INFORMÇÕES COMO LOGIN E SENHA DO USUÁRIO.
navegador.find_element(By.XPATH, '//*[@id="mat-input-0"]').send_keys("40263170000930")
navegador.find_element(By.XPATH, '//*[@id="mat-input-1"]').send_keys("09950039983")
navegador.find_element(By.XPATH, '//*[@id="mat-input-2"]').send_keys("Monica2025@")

sleep(5)
navegador.find_element(By.XPATH, '//button[@class="mat-raised-button mat-primary"]').click()

# MAXIMIZA A PÁGINA E DIMINUI O ZOOM PARA 50%
pyautogui.hotkey('winleft', 'up')
pyautogui.keyDown('ctrl')
pyautogui.hotkey('-')
pyautogui.hotkey('-')
pyautogui.hotkey('-')
pyautogui.hotkey('-')
pyautogui.hotkey('-')
pyautogui.keyUp('ctrl')
sleep(5)

hover = navegador.find_element(By.XPATH,
                               '/html/body/app-root/app-navegacao/mat-sidenav-container/mat-sidenav-content/p-menubar/div/p-menubarsub/ul/li[2]')
act = ActionChains(navegador)
act.move_to_element(hover).perform()
sleep(1)

navegador.find_element(By.XPATH,
                       '/html/body/app-root/app-navegacao/mat-sidenav-container/mat-sidenav-content/p-menubar/div/p-menubarsub/ul/li[2]/p-menubarsub/ul/li[4]/a/span').click()
sleep(1.5)

# Clicar na célula da tabela para abrir o dropdown
celula = WebDriverWait(navegador, 10).until(
    EC.element_to_be_clickable((By.XPATH, '/html/body/app-root/app-navegacao/mat-sidenav-container/mat-sidenav-content/app-meus-mtrs/mat-sidenav-container/mat-sidenav-content/p-dialog[2]/div/div[2]/form/mat-card/p-table/div/div[2]/table/tbody/tr/td[3]'))  # Substitua pelo XPath correto
)
celula.click()

# Pequeno delay para garantir que o dropdown carregue (caso precise)
sleep(1)

# **Verificar se há um iframe** e mudar o contexto (importante)
iframes = navegador.find_elements(By.TAG_NAME, "iframe")
if len(iframes) > 0:
    print(f"Encontrado {len(iframes)} iframe(s), trocando de contexto...")
    navegador.switch_to.frame(iframes[0])  # Pode ser que precise testar outros índices

# **Mover o mouse para o dropdown para garantir que ele apareça**
dropdown = WebDriverWait(navegador, 10).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="cdk-overlay-1"]/div/div'))
)
ActionChains(navegador).move_to_element(dropdown).perform()

# **Esperar os itens aparecerem e capturá-los**
itens = WebDriverWait(navegador, 10).until(
    EC.presence_of_all_elements_located((By.XPATH, '//*[@id="cdk-overlay-1"]/div/div//mat-option'))
)

# **Verificar e listar os itens**
print(f"🔍 {len(itens)} itens encontrados no dropdown:")
for item in itens:
    print(item.text)

    # Se quiser selecionar um item específico
    if item.text.strip() == "Coprocessamento":
        item.click()
        break

# **Voltar para o contexto principal (se entrou em um iframe)**
navegador.switch_to.default_content()
