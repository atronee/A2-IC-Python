from bs4 import BeautifulSoup # Biblioteca para Webscraping
import requests # Biblioteca para acessar sites reais

def estrutura(lista_acoes,lista_moedas): # Organiza os dados em um dicionário com uma lista de dicionarios para ações e para moedas
    carteira ={}
    acoes_lista = []
    for i in range(0,len(lista_acoes), 2):
        disc_acoes = {}
        disc_acoes["Nome"] = lista_acoes[i]
        disc_acoes["Quantidade"] = lista_acoes[i+1]
        acoes_lista.append(disc_acoes)
    carteira["acao"] = acoes_lista
    moedas_lista = []
    for i in range(0,len(lista_moedas), 2):
        disc_moedas = {}
        disc_moedas["Nome"] = lista_moedas[i]
        disc_moedas["Quantidade"] = lista_moedas[i+1]
        moedas_lista.append(disc_moedas)
    carteira["moeda"] = moedas_lista
    print(carteira)
    return carteira

lista_acoes = [] # Lista com as linhas dentro da tabela de ações
lista_moedas =[] # Lista com as linhas dentro da tabela de Moedas

def encontra_ativos(url):
    content = requests.get(url).text
    soup = BeautifulSoup(content, 'lxml')
    linhas_acoes = soup.find_all('tr') # Encontra todas as tags tr, que definem novas linhas
    for acao in linhas_acoes:
        todas_linhas = acao.find_all('td') # Encontra todas as tags td, linhas de tabela sem contar cabeçalho
        for linha in todas_linhas:
            if linha is not None:
                linhas_acoes = linha.find_parents("div", class_="acao") # Verifica se a tag div com classe acao é pai da tag td
                linhas_moedas = linha.find_parents("div", class_="moeda") # Verifica se a tag div com classe moeda é pai da tag td
                for items in linhas_acoes:
                    lista_acoes.append(linha.text)
                for items in linhas_moedas:
                    lista_moedas.append(linha.text)

    
encontra_ativos("https://atronee.github.io/A2-IC-Python/")
estrutura(lista_acoes,lista_moedas)