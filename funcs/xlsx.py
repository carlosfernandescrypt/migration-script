from openpyxl import load_workbook

# Carregar o arquivo Excel
wb = load_workbook('planilhas/planilha.xlsx')  # Substitua pelo caminho do seu arquivo

# Instancia
ws = wb.active

# Ler dados de columa específicas
coluna_g = [cell.value for cell in ws['g']]


print(f'Valor da célula A1: {coluna_g}')

# fecha a instancia
wb.close()
