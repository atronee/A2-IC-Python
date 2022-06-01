from openpyxl import Workbook
from openpyxl.styles import Alignment, Font


def inicializa_planilha():  # Criação de variáveis do Excel
    _planilha = Workbook()  # cria planilha
    _folha = _planilha.active  # acessa folha da planilha
    return _planilha, _folha


def formatacao_inicial(_folha):  # Indica aspectos visuais básicos da planilha, como nome e largura de colunas
    _folha.title = "Dashboard"  # altera nome da folha

    for coluna in "ABCDEFGH":  # ajuste de largura de coluna
        _folha.column_dimensions[coluna].width = 27

    for linha in range(1, 3):  # ajuste de altura de linha
        _folha.row_dimensions[linha].height = 27


def criar_titulo(_folha, ini_row, ini_col, fim_row, fim_col, nome, tamanho="12"):  # Cria uma célula de título
    celula = _folha.merge_cells(start_row=ini_row, start_column=ini_col, end_row=fim_row, end_column=fim_col)
    celula = _folha.cell(row=ini_row, column=ini_col, value=nome)
    celula.alignment = Alignment(horizontal="center", vertical="center")
    celula.font = Font(bold=True, size=tamanho)


def celulas_fixas(_folha):  # Cria as células que tem posições fixas na planilha
    # Criação de título
    criar_titulo(_folha, 1, 1, 2, 8, "Resumo da Carteira", "20")

    # Espaço para ações
    criar_titulo(_folha, 3, 1, 4, 4, "Ações", "16")
    criar_titulo(_folha, 5, 1, 5, 1, "Nome")
    criar_titulo(_folha, 5, 2, 5, 2, "Quantidade")
    criar_titulo(_folha, 5, 3, 5, 3, "Valor da ação (R$)")
    criar_titulo(_folha, 5, 4, 5, 4, "Valor acumulado (R$)")

    # Espaço para moedas
    criar_titulo(_folha, 3, 5, 4, 8, "Moedas", "16")
    criar_titulo(_folha, 5, 5, 5, 5, "Nome")
    criar_titulo(_folha, 5, 6, 5, 6, "Quantidade")
    criar_titulo(_folha, 5, 7, 5, 7, "Valor da ação (R$)")
    criar_titulo(_folha, 5, 8, 5, 8, "Valor acumulado (R$)")


def listar_acoes(_carteira):  # Extrai as ações da carteira
    acoes = []
    for ativo in _carteira["acao"]:
        acoes.append([ativo["Nome"], ativo["Quantidade"], ativo["preco_atualizado"]])

    return acoes


def apresentar_acoes(_folha, _acoes):  # Lista as ações da carteira na planilha
    num_acoes = len(_acoes)

    for i in range(num_acoes):  # Iterador das ações
        _folha.cell(row=(6 + i), column=1, value=_acoes[i][0])  # Nome da ação
        _folha.cell(row=(6 + i), column=2, value=float(_acoes[i][1]))  # Quantidade na carteira
        cotacao = _folha.cell(row=(6 + i), column=3, value=float(_acoes[i][2]))  # Valor atual da ação
        cotacao.number_format = "R$#,##0.00"  # Formato de moeda real
        contrib = _folha.cell(row=(6 + i), column=4,
                              value="=B" + str(6 + i) + "*C" + str(6 + i))  # Contribuição da ação (valor*quantidade)
        contrib.number_format = "R$#,##0.00"  # Formato de moeda real


def listar_moedas(_carteira):  # Extrai as moedas da carteira
    moedas = []
    for ativo in _carteira["moeda"]:
        moedas.append([ativo["Nome"], ativo["Quantidade"], ativo["preco_atualizado"]])

    return moedas


def apresentar_moedas(_folha, _moedas):  # Lista as moedas da carteira na planilha
    num_moedas = len(_moedas)

    # Apresentação de Moedas
    for i in range(num_moedas):  # Iterador de moedas
        _folha.cell(row=(6 + i), column=5, value=_moedas[i][0])  # Nome da moeda
        _folha.cell(row=(6 + i), column=6, value=float(_moedas[i][1]))  # Quantidade na carteira
        cotacao = _folha.cell(row=(6 + i), column=7, value=float(_moedas[i][2]))  # Valor atual da moeda
        cotacao.number_format = "R$#,##0.00"  # Formato de moeda real
        contrib = _folha.cell(row=(6 + i), column=8,
                              value="=F" + str(6 + i) + "*G" + str(6 + i))  # Contribuição da moeda (valor*quantidade)
        contrib.number_format = "R$#,##0.00"  # Formato de moeda real


def qtd_linhas(_acoes, _moedas):  # Retorna a quantidade variável de linhas
    num_acoes = len(_acoes)
    num_moedas = len(_moedas)

    return max(num_moedas, num_acoes)


def totais_acoes(_folha, _num_linhas):  # Adiciona informações do total das ações da carteira
    criar_titulo(_folha, 6 + _num_linhas, 1, 6 + _num_linhas, 1, "Total Ações")
    _folha.cell(row=(6 + _num_linhas), column=2,
                value="=SUM(B6:B" + str(5 + _num_linhas) + ")")  # Soma das quantidades de ações
    soma_val = _folha.cell(row=(6 + _num_linhas), column=3,
                           value="=SUM(C6:C" + str(5 + _num_linhas) + ")")  # Soma dos valores das ações
    soma_val.number_format = "R$#,##0.00"  # Formato de moeda real
    soma_contrib = _folha.cell(row=(6 + _num_linhas), column=4,
                               value="=SUM(D6:D" + str(5 + _num_linhas) + ")")  # Soma das contribuições das ações
    soma_contrib.number_format = "R$#,##0.00"  # Formato de moeda real


def totais_moedas(_folha, _num_linhas):  # Adiciona informações do total das moedas da carteira
    criar_titulo(_folha, 6 + _num_linhas, 5, 6 + _num_linhas, 5, "Total Moedas")
    _folha.cell(row=(6 + _num_linhas), column=6,
                value="=SUM(F6:F" + str(5 + _num_linhas) + ")")  # Soma das quantidades de moedas
    soma_val = _folha.cell(row=(6 + _num_linhas), column=7,
                           value="=SUM(G6:G" + str(5 + _num_linhas) + ")")  # Soma dos valores das ações
    soma_val.number_format = "R$#,##0.00"  # Formato de moeda real
    soma_contrib = _folha.cell(row=(6 + _num_linhas), column=8,
                               value="=SUM(H6:H" + str(5 + _num_linhas) + ")")  # Soma das contribuições das moedas
    soma_contrib.number_format = "R$#,##0.00"  # Formato de moeda real


def total_carteira(_folha, _num_linhas):  # Apresentação do Total da carteira
    criar_titulo(_folha, 9 + _num_linhas, 4, 10 + _num_linhas, 5, "Valor da Carteira", "16")
    criar_titulo(_folha, 11 + _num_linhas, 4, 11 + _num_linhas, 4, "Quantidade")
    criar_titulo(_folha, 11 + _num_linhas, 5, 11 + _num_linhas, 5, "Valor acumulado total (R$)")
    soma_val = _folha.cell(row=12 + _num_linhas, column=4, value="=B" + str(6 + _num_linhas) + "+F"
                           + str(6 + _num_linhas))  # Soma das quantidades de ações e de moedas
    soma_val.number_format = "R$#,##0.00"  # Formato de moeda real
    soma_contrib = _folha.cell(row=12 + _num_linhas, column=5, value="=D" + str(6 + _num_linhas) + "+H"
                               + str(6 + _num_linhas))  # Soma dos valores acumulados de ações e de moedas
    soma_contrib.number_format = "R$#,##0.00"  # Formato de moeda real


def salvar_excel(_planilha, nome_arquivo):  # Salva no diretório do usuário
    _planilha.save(nome_arquivo + ".xlsx")


def dashboard(_carteira, _nome):  # Consolidação do módulo; cria dashboard com dados da carteira
    planilha, folha = inicializa_planilha()

    formatacao_inicial(folha)

    celulas_fixas(folha)

    acoes = listar_acoes(_carteira)
    moedas = listar_moedas(_carteira)

    apresentar_acoes(folha, acoes)
    apresentar_moedas(folha, moedas)

    num_linhas = qtd_linhas(acoes, moedas)

    totais_acoes(folha, num_linhas)
    totais_moedas(folha, num_linhas)

    total_carteira(folha, num_linhas)

    salvar_excel(planilha, _nome)
