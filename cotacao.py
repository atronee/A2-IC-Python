from datetime import datetime
import yfinance as yf
import pandas as pd

#Site: https://finance.yahoo.com/

#Carteira Exemplo
carteira = {"acao":[{"Nome":"PETR4.SA", "Quantidade":"1000"},{"Nome":"VALE3.SA", "Quantidade":"1000"}], 
            "moeda": [{"Nome":"USDBRL=X", "Quantidade":"100"}, {"Nome":"EURBRL=X", "Quantidade":"500"}]}

for tipo, ativos in carteira.items():
    for ativo in ativos:
        ticker = yf.Ticker(ativo["Nome"])
        cotacao_do_dia = ticker.history(period="1d")
        ativo["preco_atualizado"] = cotacao_do_dia.iloc[0]["Close"] # Coisa do pandas
        ativo["preço_histórico"] = ticker.history(period="1y")   # tipo <class 'pandas.core.frame.DataFrame'>

print(carteira)
