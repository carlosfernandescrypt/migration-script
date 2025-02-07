from openpyxl import load_workbook

# Caminho da planilha
caminho_planilha = "funcs/planilha.xlsx"

# Carregar planilha
def carregar_planilha():
    """Carrega a planilha e retorna a planilha e a aba ativa."""
    try:
        wb = load_workbook(caminho_planilha)
        ws = wb.active
        return wb, ws
    except Exception as e:
        print(f"Erro ao carregar a planilha: {str(e)}")
        return None, None

def verificar_oculto(ws):
    """Verifica na coluna H da planilha de controle se a página deve ser oculta."""
    try:
        valor_h4 = ws["H8"].value
        if valor_h4 == "Oculta":
            print("A página deve ser oculta.")
            return True
        else:
            print("A página não deve ser oculta.")
            return False
    except Exception as e:
        print(f"Erro ao verificar se a página deve ser oculta: {str(e)}")
        return False

def pegar_conteudo_input_por_tag_e_class(ws):
    """Obtém o conteúdo de um campo de entrada (input) usando tags e classes da planilha."""
    try:
        # Localiza o valor de uma célula específica na planilha, por exemplo, P4
        valor = ws["P4"].value
        print(f"Valor da célula P4: {valor}")
        return valor
    except Exception as e:
        print(f"Erro ao pegar o conteúdo da célula: {str(e)}")
        return None

def pegar_conteudo_input_por_id(ws):
    """Obtém o conteúdo de um campo de entrada (input) pelo ID da planilha."""
    try:
        # Localiza o valor de uma célula específica na planilha, por exemplo, P4
        valor = ws["P4"].value
        print(f"Valor da célula P4: {valor}")
        return valor
    except Exception as e:
        print(f"Erro ao pegar o conteúdo da célula: {str(e)}")
        return None
