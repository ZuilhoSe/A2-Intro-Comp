import yfinance as yf

dic = {"MGLU3.SA":"1000", "PETR4.SA":"100000"}
#A função recebe um dicionario que contém o nome das açoes como chaves e retorna a cotação do ativo na última semana útil
def cotacao_semana(dic):
    lista_semanal = []
    for x in dic.keys():
        y=yf.Ticker(x)
        y_hist = y.history(period="5d")
        lista_semanal.append(y_hist)
    return lista_semanal

print(cotacao_semana(dic)[0])

#A função recebe um dicionario que contém o nome das açoes como chaves e retorna a cotação do ativo no último ano
def cotacao_anual(dic):
    lista_anual = []
    for x in dic.keys():
        y=yf.Ticker(x)
        y_hist = y.history(period="1y")
        lista_anual.append(y_hist)
    return lista_anual

print(cotacao_anual(dic)[0])