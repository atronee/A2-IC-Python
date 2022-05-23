from bs4 import BeautifulSoup # Biblioteca para Webscraping
import requests # Biblioteca para acessar sites reais

def encontra_ativos(url):
    with open(url, 'r') as html_file:
        content = html_file.read()
        #html_text = requests.get(url).text  Real projeto abrir um url
        soup = BeautifulSoup(content, 'lxml')
        ativos = soup.find_all('tr', class_="linha")
        dicionário = {}
        for index, ativo in enumerate(ativos):   # Enumera cada tag "tr"
            nome_compania = ativo.find('td', class_='compania').text.replace(' ','')
            ações = ativo.find('td', class_='nomeacao').text.replace(' ','')
            num_ações = ativo.find('td', class_='quantidadeacoes').text.replace(' ','')
            dicionário[index] = {"nome da compania/ Pais":nome_compania, 
                                 "nome do Ativo":ações,
                                 "Quantidade":num_ações}
    return dicionário
    # Cria um dicionário com chave para cada linha da tabela
