import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import load_workbook
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys

# Configurações
url_base = "http://localhost:8080"
caminho_planilha = "./planilhas/planilha.xlsx"

# Inicialização do WebDriver
service = Service(ChromeDriverManager().install())
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--no-sandbox")
driver = webdriver.Chrome(service=service, options=options)
wait = WebDriverWait(driver, 20)

# Carregar planilha
wb = load_workbook(caminho_planilha)
ws = wb.active

def fazer_login():
    """Realiza o login no sistema."""
    try:
        print("Acessando a página de login...")
        driver.get(url_base)
        
        # Clicar no botão de login
        print("Clicando no botão de login...")
        botao_login = wait.until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, ".sign-in.text-default.btn.btn-sm.btn-unstyled")
        ))
        botao_login.click()
        
        # Preencher e-mail
        print("Preenchendo e-mail...")
        campo_email = wait.until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, ".field.clearable.form-control")
        ))
        campo_email.clear()
        campo_email.send_keys("test@liferay.com")
        
        # Preencher senha
        print("Preenchendo senha...")
        campo_senha = wait.until(EC.element_to_be_clickable(
            (By.ID, "_com_liferay_login_web_portlet_LoginPortlet_password")
        ))
        campo_senha.clear()
        campo_senha.send_keys("admin")
        
        # Clicar no botão de entrar
        print("Clicando no botão de entrar...")
        botao_entrar = wait.until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, ".btn.btn-primary")
        ))
        botao_entrar.click()


        # Navegar para a área de administração usando a nova classe
        print("Navegando para área de administração...")
        wait.until(EC.element_to_be_clickable(
            (By.ID, "panel-manage-site_administration_build-link")
        )).click()
        
        # Clicar no link de páginas
        print("Acessando seção de páginas...")
        wait.until(EC.element_to_be_clickable(
            (By.ID, "_com_liferay_product_navigation_product_menu_web_portlet_ProductMenuPortlet_portlet_com_liferay_layout_admin_web_portlet_GroupPagesPortlet")
        )).click()
        
        return True
    except Exception as e:
        print(f"Erro ao fazer login: {str(e)}")
        return False


# Verificar se o localhost está acessível
if not fazer_login():
    print(f"Erro: Não foi possível fazer login em {url_base}. Verifique as credenciais e o servidor.")
    driver.quit()
    exit()




def clicar_botao(css_selector):
    """Clica em um botão específico com base no seletor CSS."""
    try:
        botao = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, css_selector)))
        botao.click()
        time.sleep(1)  # Pequeno delay para garantir a transição
    except Exception as e:
        print(f"Erro ao clicar no botão {css_selector}: {str(e)}")

def clicar_botao_novo():
    """Clica no botão 'Novo' usando seu texto específico."""
    try:
        botao_novo = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@aria-label='Novo' and @title='Novo']//span[text()='Novo']")))
        botao_novo.click()
        time.sleep(1)
    except Exception as e:
        print(f"Erro ao clicar no botão 'Novo': {str(e)}")

def clicar_link_pagina():
    """Clica no link 'Página' dentro do menu suspenso."""
    try:
        link_pagina = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[text()='Página']")))
        link_pagina.click()
        time.sleep(1)
    except Exception as e:
        print(f"Erro ao clicar no link 'Página': {str(e)}")

def clicar_div_pagina_definida():
    """Clica na div 'Página definida'."""
    try:
        div_pagina_definida = wait.until(EC.element_to_be_clickable((By.XPATH, "//p[@class='card-title' and @title='Página definida']//span[text()='Página definida']")))
        div_pagina_definida.click()
        time.sleep(1)
    except Exception as e:
        print(f"Erro ao clicar na div 'Página definida': {str(e)}")

def clicar_div_pagina_widget():
    """Clica na div 'Página Widget'."""
    try:
        div_pagina_widget = wait.until(EC.element_to_be_clickable((By.XPATH, "//p[@class='card-title' and @title='Página de Widget']//span[text()='Página de Widget']")))
        div_pagina_widget.click()
        time.sleep(1)
    except Exception as e:
        print(f"Erro ao clicar na div 'Página definida': {str(e)}")

def clicar_div_vincular_pagina_deste_site():
    """Clica na div 'Vincular Pagina Deste Site'."""
    try:
        clicar_div_vincular_pagina_deste_site = wait.until(EC.element_to_be_clickable((By.XPATH, "//p[@class='card-title' and @title='Vincular a uma página deste site']//span[text()='Vincular a uma página deste site']")))
        clicar_div_vincular_pagina_deste_site.click()
        time.sleep(1)
    except Exception as e:
        print(f"Erro ao clicar na div 'Página definid': {str(e)}")

def clicar_div_vincular_url():
    """Clica na div 'Vincular Url'."""
    try:
        div_vincular_url = wait.until(EC.element_to_be_clickable((By.XPATH, "//p[@class='card-title' and @title='Vincular a uma URL']//span[text()='Vincular a uma URL']")))
        div_vincular_url.click()
        time.sleep(1)
    except Exception as e:
        print(f"Erro ao clicar na div 'Página definida': {str(e)}")

    def preencher_input_nome():
        """Muda para o iframe, espera 1 segundo, clica no campo de nome, preenche com 'Teste' e pressiona Enter."""
        try:
            time.sleep(1)  # Espera antes de executar a função

            iframe = wait.until(EC.presence_of_element_located((By.ID, "addLayoutDialog_iframe_")))

            print("[LOG] Alternando para o iframe...")
            driver.switch_to.frame(iframe)  # Mudar para dentro do iframe

            campo_nome = wait.until(EC.element_to_be_clickable(
                (By.ID, "_com_liferay_layout_admin_web_portlet_GroupPagesPortlet_name")
            ))

            campo_nome.click()  # Garante que o campo seja ativado

            campo_nome.clear()  # Remove qualquer texto anterior

            campo_nome.send_keys("Teste")  # Digita o texto

            campo_nome.send_keys(Keys.RETURN)  # Pressiona Enter para confirmar

            # Retornar ao contexto principal
            driver.switch_to.default_content()

        except Exception as e:
            print(f"[ERRO] Falha ao preencher o campo de nome dentro do iframe: {str(e)}")
            driver.switch_to.default_content()  # Retorna ao contexto principal mesmo em caso de erro


def apertar_enter(campo_input):
    """Simula o pressionamento da tecla Enter no campo de entrada."""
    try:
        campo_input.send_keys(Keys.RETURN)
        print("Tecla Enter pressionada.")
    except Exception as e:
        print(f"Erro ao pressionar Enter: {str(e)}")

def criar_pagina(type):
    if type == "Definida":
        clicar_div_pagina_definida()
        preencher_input_nome()
    elif type == "Widget":
        clicar_div_pagina_widget()
        preencher_input_nome()
    elif type == "Vincular a uma página deste site":
        clicar_div_vincular_pagina_deste_site()
        preencher_input_nome()
    elif type == "Vincular a uma URL":
        clicar_div_vincular_url()
        preencher_input_nome()
    pass

try:

    # clicar no botão class="dropdown-toggle nav-btn btn btn-primary"
    clicar_botao_novo()
    # class="dropdown-item" Página
    clicar_link_pagina()
    valor_p4 = ws["P4"].value
    valor_extraido = valor_p4.split(": ", 1)[-1]

    print(f"Valor da célula P4: {valor_extraido}")

    criar_pagina(valor_extraido)

finally:
    wb.close()
    print("Planilha fechada.")