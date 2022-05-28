from bs4 import BeautifulSoup # Biblioteca para Webscraping
import requests # Biblioteca para acessar sites reais

def estrutura(lista_acoes,lista_moedas, carteira): # Organiza os dados em um dicionário com uma lista de dicionarios para ações e para moedas
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
    return carteira

lista_acoes = [] # Lista com as linhas dentro da tabela de ações
lista_moedas =[] # Lista com as linhas dentro da tabela de Moedas
carteira ={}

def encontra_ativos(url):
    content = requests.get(url).text
    soup = BeautifulSoup(content, 'lxml')
    todas_linhas = soup.find_all('tr') # Encontra todas as tags tr, que definem novas linhas
    for linha in todas_linhas:
        todas_celulas = linha.find_all('td') # Encontra todas as tags td, celulas da tabela sem contar cabeçalho
        for celula in todas_celulas:
            if celula is not None:
                celulas_acoes = celula.find_parents("div", class_="acao") # Verifica se a tag div com classe acao é pai da tag td
                celulas_moedas = celula.find_parents("div", class_="moeda") # Verifica se a tag div com classe moeda é pai da tag td
                for i in celulas_acoes:
                    lista_acoes.append(celula.text)
                for j in celulas_moedas:
                    lista_moedas.append(celula.text)

    estrutura(lista_acoes,lista_moedas, carteira)
    