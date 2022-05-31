import yfinance as yf

#Site: https://finance.yahoo.com/

#Carteira Exemplo
carteira = {"acao":[{"Nome":"AAPL34.SA", "Quantidade":"15"},{"Nome":"MSFT", "Quantidade":"18"}, {"Nome":"NVDA", "Quantidade":"19"}, {"Nome":"INTC", "Quantidade":"12"}, {"Nome":"AMD", "Quantidade":"34"}], 
           "moeda": [{"Nome":"EURBRL=X", "Quantidade":"100"}, {"Nome":"USDBRL=X", "Quantidade":"500"}]}

def cotacao(carteira):
    for ativos in carteira.values():
        for ativo in ativos:
            ticker = yf.Ticker(ativo["Nome"])
            ticker_info = ticker.info
            preco_do_dia = ticker_info["regularMarketPrice"] # float
            cotacao_do_ano = ticker.history(period="1y", interval="1d") #Dataframe com o histórico do ativo
            preco_historico = cotacao_do_ano.loc[:,['Close']] #Dataframe com o histórico do fechamento do ativo 
            
            moeda_do_ativo = ticker_info["currency"] # String

            if moeda_do_ativo != "BRL":
                cotacao_br = converte_moeda_de_cotacao(moeda_do_ativo, preco_historico)
                preco_do_dia = preco_do_dia * cotacao_br["preco_atualizado"]
                preco_historico = preco_historico * cotacao_br["preco_historico"]

            ativo["preco_atualizado"] = preco_do_dia 
            ativo["preco_historico"] = preco_historico 
    return carteira

def converte_moeda_de_cotacao(moeda_estrangeira, cotacao_ativo):
    moeda_br = moeda_estrangeira + "BRL=X"
    convecao_moeda = {}
    convecao_moeda["preco_atualizado"] = yf.Ticker(moeda_br).info["regularMarketPrice"]
    convecao_moeda["preco_historico"] = yf.Ticker(moeda_br).history(start=cotacao_ativo.index[0], stop=cotacao_ativo.index[-1], interval="1d").loc[:,['Close']]
    return convecao_moeda

cota = cotacao(carteira)
print(cota)