from bs4 import BeautifulSoup
import requests

def estrutura(lista_acoes,lista_moedas, carteira):
    """-> Organiza os dados em um dicionário com informações das ações e moedas.

    Args:
        :lista_acoes(list): A lista que contem as informações de uma ação em pares ex: [nome, petr4, quantidade, 100].\n
        :lista_moedas(list): A lista que contem as informações de uma moeda em pares.\n
        :carteira(dict): o dicionário onde seram adicionadas as informações.
    Returns:
        dict: retornará um dicionário com a estrutura {acao:[{nome:"nome",quantidade:"número"}],\n
                                                      moeda:[{nome:"nome",quantidade:"número"}],\n
                                                      título:"título"}.
    """
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

def encontra_ativos(url):
    """-> Busca em um html, as celulas com div de class acao ou moeda e adiciona as informações numa lista em pares.

    Args:
        :url(string): o url de uma carteira.

    Returns:
        dict: retornará a função estrutura com lista_moedas, lista_acoes e o dicionário carteira como parâmetro.
    """   
    lista_acoes = [] # Lista com as células dentro da tabela de ações
    lista_moedas =[] # Lista com as células dentro da tabela de Moedas
    carteira ={}     # Dicionário que vai ser retornado para a interface
    content = requests.get(url).text 
    soup = BeautifulSoup(content, 'lxml')
    todas_linhas = soup.find_all('tr')        # Encontra todas as tags tr, que definem novas linhas
    for linha in todas_linhas:
        todas_celulas = linha.find_all('td')  # Encontra todas as tags td, celulas da tabela sem contar cabeçalho
        for celula in todas_celulas:
            if celula is not None:                                          # A lógica é se uma célula não é nula, o programa vai
                celulas_acoes = celula.find_parents("div", class_="acao")   # Verificar se a tag div com classe acao é pai da tag td,
                celulas_moedas = celula.find_parents("div", class_="moeda") # Verifica se a tag div com classe moeda é pai da tag td

                for i in celulas_acoes:                                     # As iterações para adicionar cada célula em sua lista
                    lista_acoes.append(celula.text)
                for j in celulas_moedas:
                    lista_moedas.append(celula.text)

    return estrutura(lista_acoes,lista_moedas, carteira)