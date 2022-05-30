from openpyxl import Workbook
from openpyxl.styles import Alignment, Font

def criar_titulo(ini_row, ini_col, fim_row, fim_col, nome, tamanho="12"):  # Cria uma célula de título
    celula = ws.merge_cells(start_row=ini_row, start_column=ini_col, end_row=fim_row, end_column=fim_col)
    celula = ws.cell(row=ini_row, column=ini_col, value=nome)
    celula.alignment = Alignment(horizontal="center", vertical="center")
    celula.font = Font(bold=True, size=tamanho)

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

wb = Workbook()  # cria planilha
ws = wb.active  # acessa folha da planilha
ws.title = "Dashboard"  # altera nome da folha
for i in "ABCDEFGH": # Ajuste de largura
    ws.column_dimensions[i].width = 27
for i in range(1, 3): # Ajuste de altura
    ws.row_dimensions[i].height = 27

# Criação de título
criar_titulo(1, 1, 2, 8, "Resumo da Carteira", "20")

# Espaço para ações
criar_titulo(3, 1, 4, 4, "Ações", "16")
criar_titulo(5, 1, 5, 1, "Nome")
criar_titulo(5, 2, 5, 2, "Quantidade")
criar_titulo(5, 3, 5, 3, "Valor da ação (R$)")
criar_titulo(5, 4, 5, 4, "Valor acumulado (R$)")

# Espaço para moedas
criar_titulo(3, 5, 4, 8, "Moedas", "16")
criar_titulo(5, 5, 5, 5, "Nome")
criar_titulo(5, 6, 5, 6, "Quantidade")
criar_titulo(5, 7, 5, 7, "Valor da ação (R$)")
criar_titulo(5, 8, 5, 8, "Valor acumulado (R$)")

acoes = list(carteira["Ações"].keys())
moedas = list(carteira["Moedas"].keys())
num_acoes = len(acoes)
num_moedas = len(moedas)

# Apresentação de Ações
for i in range(num_acoes):  # Iterador das ações
    ws.cell(row=(6 + i), column=1, value=acoes[i])  # Nome da ação
    ws.cell(row=(6 + i), column=2, value=carteira["Ações"][acoes[i]])  # Quantidade na carteira
    cotacao = ws.cell(row=(6 + i), column=3, value=cota["Ações"][acoes[i]])  # Valor atual da ação
    cotacao.number_format = "R$#,##0.00"  # Formato de moeda real
    contrib = ws.cell(row=(6 + i), column=4, value="=B" + str(6 + i) + "*C" + str(6 + i))  # Contribuição da ação (valor*quantidade)
    contrib.number_format = "R$#,##0.00" # Formato de moeda real

# Apresentação de Moedas
for i in range(num_moedas):  # Iterador de moedas
    ws.cell(row=(6 + i), column=5, value=moedas[i])  # Nome da moeda
    ws.cell(row=(6 + i), column=6, value=carteira["Moedas"][moedas[i]])  # Quantidade na carteira
    cotacao = ws.cell(row=(6 + i), column=7, value=cota["Moedas"][moedas[i]])  # Valor atual da moeda
    cotacao.number_format = "R$#,##0.00" # Formato de moeda real
    contrib = ws.cell(row=(6 + i), column=8, value="=F" + str(6 + i) + "*G" + str(6 + i))  # Contribuição da moeda (valor*quantidade)
    contrib.number_format = "R$#,##0.00" # Formato de moeda real

# Apresentação dos totais
num_linhas = max(num_acoes, num_moedas)

criar_titulo(6 + num_linhas, 1, 6 + num_linhas, 1, "Total Ações")
ws.cell(row=(6 + num_linhas), column=2, value="=SUM(B6:B" + str(5 + num_linhas) + ")")  # Soma das quantidades de ações
soma_val = ws.cell(row=(6 + num_linhas), column=3, value="=SUM(C6:C" + str(5 + num_linhas) + ")")  # Soma dos valores das ações
soma_val.number_format = "R$#,##0.00" # Formato de moeda real
soma_contrib = ws.cell(row=(6 + num_linhas), column=4, value="=SUM(D6:D" + str(5 + num_linhas) + ")")  # Soma das contribuições das ações
soma_contrib.number_format = "R$#,##0.00" # Formato de moeda real

criar_titulo(6 + num_linhas, 5, 6 + num_linhas, 5, "Total Moedas")
ws.cell(row=(6 + num_linhas), column=6, value="=SUM(F6:F" + str(5 + num_linhas) + ")")  # Soma das quantidades de moedas
soma_val = ws.cell(row=(6 + num_linhas), column=7, value="=SUM(G6:G" + str(5 + num_linhas) + ")")  # Soma dos valores das ações
soma_val.number_format = "R$#,##0.00" # Formato de moeda real
soma_contrib = ws.cell(row=(6 + num_linhas), column=8, value="=SUM(H6:H" + str(5 + num_linhas) + ")")  # Soma das contribuições das moedas
soma_contrib.number_format = "R$#,##0.00" # Formato de moeda real

# Apresentação do Total da carteira
criar_titulo(9 + num_linhas, 4, 10 + num_linhas, 5, "Valor da Carteira", "16")
criar_titulo(11 + num_linhas, 4, 11 + num_linhas, 4, "Quantidade")
criar_titulo(11 + num_linhas, 5, 11 + num_linhas, 5, "Valor acumulado total (R$)")
soma_val = ws.cell(row=12 + num_linhas, column=4, value="=B" + str(6 + num_linhas) + "+F" + str(6 + num_linhas))  # Soma das quantidades de ações e de moedas
soma_val.number_format = "R$#,##0.00" # Formato de moeda real
soma_contrib = ws.cell(row=12 + num_linhas, column=5, value="=D" + str(6 + num_linhas) + "+H" + str(6 + num_linhas))  # Soma dos valores acumulados de ações e de moedas
soma_contrib.number_format = "R$#,##0.00" # Formato de moeda real

# Salvando
wb.save("Carteira Cláudia.xlsx")
