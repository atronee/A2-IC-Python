from bs4 import BeautifulSoup # Biblioteca para Webscraping
import requests # Biblioteca para acessar sites reais

def encontra_ativos(url):
    with open(url, 'r') as html_file:
        content = html_file.read()
        #html_text = requests.get(url).text  Real projeto abrir um url
        soup = BeautifulSoup(content, 'lxml')
        carteira = {} # Um Dicionário que contém todas as informações de ações e moedas
        dicionario_acoes = {} # Dicionário com as informações de uma ação
        dicionario_moedas = {} #Dicionário com as informações de uma moeda
        lista_acoes = [] # Lista com os Dicionários de cada ação
        lista_moedas = [] # Lista com os Dicionários de cada moeda
        acoes = soup.find_all('div', class_="acao") # Encontra todas as tags div com classe "acao"
        for index, ativo in enumerate(acoes):   # Enumera cada tag "tr"
            nome_acao = ativo.find('td').text.replace(' ','')
            num_acoes = ativo.find('td').text.replace(' ','')
            dicionario_acoes[index] = {"Ação":nome_acao,
                                       "Quantidade":num_acoes}
            lista_acoes.append(dicionario_acoes) # Adiciona o dicionário de cada ação a lista de ações
        moedas = soup.find_all('div', class_="moeda") # Encontra todas as tags div com classe "moeda"
        for index, ativo in enumerate(moedas):   # Enumera cada tag "tr"
            nome_acao = ativo.find('td').text.replace(' ','')
            num_acoes = ativo.find('td').text.replace(' ','')
            dicionario_moedas[index] = {"Moeda":nome_acao,
                                        "Quantidade":num_acoes}
            lista_moedas.append(dicionario_moedas) # Adiciona o dicionário de cada moeda a lista de moedas
        carteira["acao"] = lista_acoes 
        carteira["moeda"] = lista_moedas
    return carteira
    # Cria um dicionário com chave para cada linha da tabela
