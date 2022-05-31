import yfinance as yf

def cotacao(carteira):
    """
    -> Recebe uma carteira com ativos e retorna a cotaçao do dia e a cotação histórica.
    :param carteira: carteira com os ativos. 
    :return: retorna a carteira com a cotação atualizada do dia e a cotação histórica.
    """
    
    for ativos in carteira.values():
        for ativo in ativos:
            ticker = yf.Ticker(ativo["Nome"]) 
            ticker_info = ticker.info
            preco_do_dia = ticker_info["regularMarketPrice"] # Pega a cotaçao do momento (float)
            cotacao_do_ano = ticker.history(period="1y", interval="1d") # Pega a cotação histórica (Dataframe com o histórico do ativo)
            preco_historico = cotacao_do_ano.loc[:,['Close']] #Dataframe com o histórico do fechamento do ativo 
            
            moeda_do_ativo = ticker_info["currency"] # A moeda que o ativo está sendo cotada (String)

            if moeda_do_ativo != "BRL": #Caso a moeda não seja real, o programa irá converter os valores da moeda estrangeira para o real
                cotacao_br = converte_moeda_de_cotacao(moeda_do_ativo, preco_historico)
                preco_do_dia = preco_do_dia * cotacao_br["preco_atualizado"]
                preco_historico = preco_historico * cotacao_br["preco_historico"]

            #Adiciona na carteira 
            ativo["preco_atualizado"] = preco_do_dia 
            ativo["preco_historico"] = preco_historico 
    return carteira

def converte_moeda_de_cotacao(moeda_estrangeira, cotacao_ativo):
    """
    -> Recebe a moeda estrangeira e converte ela para a moeda Real, sendo uma cotaçao instantânea para o preço atualizado do dia e outra com a cotação do fechamento de cada dia para o preço histórico do ativo.
    :param moeda_estrangeira: moeda que se deseja fazer converter. 
    :param cotacao_ativo: o ativo que se deseja ter a cotação. 
    :return: Um dicionário com a cotação instantânea para o preço atualizado do dia e com a cotação do fechamento de cada dia para o preço histórico.
    """

    moeda_br = moeda_estrangeira + "BRL=X"
    convecao_moeda = {}
    convecao_moeda["preco_atualizado"] = yf.Ticker(moeda_br).info["regularMarketPrice"] # Cotação para o preço atualizado do dia
    convecao_moeda["preco_historico"] = yf.Ticker(moeda_br).history(start=cotacao_ativo.index[0], stop=cotacao_ativo.index[-1], interval="1d").loc[:,['Close']] #Cotação para o histórico do ativo.
    return convecao_moeda
