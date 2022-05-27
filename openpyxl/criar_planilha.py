from openpyxl import Workbook
from openpyxl.styles import Alignment, Font

# Dados esperados:

# carteira: dict de ações com qtds de papers, e de moedas, com qtds
carteira = {
    "Ações":
        {"VALE3": 1000, "MGLU3": 1000, "ITUB4": 375},
    "Moedas":
        {"CAD": 150, "CHF": 500}
}

# cota: informações esperadas do YFINANCE, como preço atual e série histórica
cota = {
    "Ações":
        {"VALE3": 84.25, "MGLU3": 4.07, "ITUB4": 26.02},
    "Moedas":
        {"CAD": 3.74, "CHF": 4.97},
    "Série Histórica": []
}

wb = Workbook()         # cria planilha
ws = wb.active          # acessa folha da planilha
ws.title = "Dashboard"  # altera nome da folha

# Criação de título
ws.merge_cells("B2:E3")
titulo = ws.cell(row = 2, column = 2, value = "Resumo da Carteira")
titulo.alignment = Alignment(horizontal = "center", vertical = "center")
titulo.font = Font(bold = True)

# Espaço para ações
ws.merge_cells("B4:C4")
ws.cell(row = 4, column = 2, value = "Ações")

# Espaço para moedas
ws.merge_cells("D4:E4")
ws.cell(row = 4, column = 4, value = "Moedas")

acoes = list(carteira["Ações"].keys())
moedas = list(carteira["Moedas"].keys())
num_acoes  = len(acoes)
num_moedas = len(moedas)

# Apresentação de Ações
for i in range(num_acoes):
    ws.cell(row = (5 + i), column = 2, value = acoes[i])
    ws.cell(row = (5 + i), column = 3, value = carteira["Ações"][acoes[i]])

ws.cell(row = (5 + num_acoes), column = 2, value = "Total Ações")
ws.cell(row = (5 + num_acoes), column = 3, value = "=SUM(C5:C" + str(4 + num_acoes) + ")")

# Apresentação de Moedas
for i in range(num_moedas):
    ws.cell(row = (5 + i), column = 4, value = moedas[i])
    ws.cell(row = (5 + i), column = 5, value = carteira["Moedas"][moedas[i]])

ws.cell(row = (5 + num_moedas), column = 4, value = "Total Moedas")
ws.cell(row = (5 + num_moedas), column = 5, value = "=SUM(E5:E" + str(4 + num_moedas) + ")")

# Apresentação Total
ultima_linha = max(6 + num_acoes, 6 + num_moedas)
ws.merge_cells("B" + str(ultima_linha) + ":E" + str(ultima_linha))
ws.cell(row = ultima_linha, column = 2, value = "TOTAL:").alignment = Alignment(horizontal = "center", vertical = "center")
ws.cell(row = ultima_linha, column = 6, value = "=SUM(C5:C" + str(4 + num_acoes) + ") + SUM(E5:E" + str(4 + num_moedas) + ")")

# Salvando
wb.save("Carteira Cláudia.xlsx")
