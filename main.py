import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from oauth2client.service_account import ServiceAccountCredentials
import gspread

# Configurações iniciais
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
credenciais = ServiceAccountCredentials.from_json_keyfile_name('credenciais.json', scope)
gc = gspread.authorize(credenciais)
planilha = gc.open_by_url('https://docs.google.com/spreadsheets/d/1KbyZSeG7_BQQN4xcYlgMQNjCSK5Dz6Oxv2Tr8tIhP80/edit?usp=sharing')
worksheet = planilha.sheet1

driver = webdriver.Chrome()
wait = WebDriverWait(driver, 20)
url_base = "http://localhost:8080"

def processar_linha(row):
    try:
        # Pular linhas vazias
        if not worksheet.cell(row, 7).value:  # Coluna G
            return True

        # Coletar dados da planilha
        tipo_pagina = worksheet.cell(row, 15).value  # Coluna O
        nome_completo = worksheet.cell(row, 7).value  # Coluna G
        visibilidade = worksheet.cell(row, 8).value   # Coluna H
        url_redirect = worksheet.cell(row, 17).value  # Coluna Q
        pagina_destino = worksheet.cell(row, 16).value  # Coluna P

        # Passos comuns a todos os casos
        def passos_comuns():
            # Abrir dropdown e selecionar item
            wait.until(EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "button.dropdown-toggle.nav-btn.btn.btn-primary")
            )).click()
            
            wait.until(EC.element_to_be_clickable(
                (By.CLASS_NAME, "dropdown-item")
            )).click()

            # Preencher nome da página
            nome_pagina = nome_completo.split('>')[-1].strip()
            campo_nome = wait.until(EC.element_to_be_clickable(
                (By.CSS_SELECTOR, ".field.form-control")
            ))
            campo_nome.clear()
            campo_nome.send_keys(nome_pagina)

            # Primeiro salvamento
            driver.find_element(By.CSS_SELECTOR, ".btn.btn-primary").click()
            time.sleep(1)

            # Verificar visibilidade
            if visibilidade == "Oculta":
                checkbox = wait.until(EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, "input[type='checkbox'].field")
                ))
                if not checkbox.is_selected():
                    checkbox.click()

        # Executar passos comuns
        passos_comuns()

        # Case 1: Vincular URL
        if tipo_pagina == "Vincular URLS":
            campo_url = wait.until(EC.element_to_be_clickable(
                (By.CSS_SELECTOR, ".field.lfr-input-text-container.form-control.error-field")
            ))
            campo_url.clear()
            campo_url.send_keys(url_redirect)

        # Case 2: Vincular a página do site
        elif tipo_pagina == "Vincular a uma página deste site":
            # Abrir seletor de páginas
            wait.until(EC.element_to_be_clickable(
                (By.CSS_SELECTOR, ".btn.btn-secondary")
            )).click()

            # Selecionar página pai
            paginas_pai = wait.until(EC.presence_of_all_elements_located(
                (By.CSS_SELECTOR, ".c-inner")
            ))
            for pagina in paginas_pai:
                if pagina.text.strip() == nome_completo.split('>')[0].strip():
                    pagina.click()
                    break

            # Selecionar subpágina
            subpaginas = wait.until(EC.presence_of_all_elements_located(
                (By.CSS_SELECTOR, ".c-inner")
            ))
            for subpagina in subpaginas:
                if subpagina.text.strip() == pagina_destino:
                    subpagina.click()
                    break

        # Passos finais comuns
        # Capturar URL amigável
        url_amigavel = wait.until(EC.presence_of_element_located(
            (By.CSS_SELECTOR, ".form-control.language-value")
        )).get_attribute('value')
        
        # Atualizar planilha
        worksheet.update_cell(row, 2, url_amigavel)  # Coluna B

        # Salvamento final
        driver.find_element(By.CSS_SELECTOR, ".btn.btn-primary").click()
        time.sleep(1)
        return True

    except Exception as e:
        print(f"Erro na linha {row}: {str(e)}")
        return False

# Processar a partir da linha 3
row = 3
while worksheet.cell(row, 7).value:  # Enquanto houver dados na coluna G
    driver.get(url_base)
    if processar_linha(row):
        print(f"Linha {row} processada com sucesso")
        row += 1  # Próxima linha somente após sucesso
    else:
        print(f"Repetindo tentativa na linha {row}")
    time.sleep(2)

driver.quit()