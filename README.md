<h1 align="center">PROJETO PY.BANK - A2 IC</h1>
Para executar o programa, é necessário rodar o arquivo main.py da pasta. Nele há apenas um código que chamará a função interface do módulo.

## Objetivo:
O programa receberá um url com uma carteira e entregará um arquivo XLSX na pasta do programa com os nomes dos ativos (ações e moedas), cotação atual, cotação histórica e o valor acumulado desse ativo. Além disso, o arquivo XLSX terá três gráficos para avaliação da qualidade da carteira e um QR code com o valor total da carteira. 

Carteira do grupo:

https://atronee.github.io/A2-IC-Python/Carteira%20Fernando


## Projeto:

### interface:

Responsável pela interface no terminal e a estrutura do programa. A interface possui duas opções: uma para adicionar o url e a outra para sair do programa. 

Ao clicar na opção de adicionar o url, a interface chamará os demais módulos do programa.

### scrapping:

Módulo responsável por acessar um url que contém uma carteira, buscar no html, as tags div com classe "acao" ou "moeda" e retornar o conteúdo em um dicionário de ações e moedas.

Este módulo utiliza Beautiful Soup, requests e lxml (parseador de html).

### cotacao:

Módulo responsável por pegar um dicionário gerado pelo scrapping e adicionar a este a cotação atual e o histórico do ativo de acordo com o site Yahoo Finance. 

Se o ativo estiver sendo cotado em uma moeda estrangeira, o módulo fará a conversão para o Real, tanto da cotação atual como da cotação histórica, utilizando um histórico da conversão para isso.

Este módulo utiliza yfinance.

### dashboard:

Módulo responsável pela criação da tabela Excel com o valor e a quantidade de cada ativo da carteira de investimentos, pela criação dos gráficos relacionados aos ativos da mesma carteira e pela criação do QRcode com o valor total da carteira. 

Este módulo utiliza openpyxl, datetime, qrcode e Image.
