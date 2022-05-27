import yfinance as yf

dic = {"MGLU3.SA":"1000", "PETR4.SA":"100000"}

def cotacao(dic):
    lista = []
    for x in dic.keys():
        y=yf.Ticker(x)
        y_hist = y.history(period="3mo", interval="1d")
        lista.append(y_hist)
    return lista

print(cotacao(dic)[0])

