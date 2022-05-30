import yfinance as yf
import time

dic = {"PETR4.SA":"10", "AMZN":"100"}

fator_conversao = 2

def cotacao_semana(dic):
    """A função recebe um dicionario que contém o nome das açoes como chaves e retorna a cotação do ativo na última semana útil
    Args:
        dic (Dictionary): Espera um dicionário onde as chaves são os códigos de cada ativo
    Returns:
        Dictionary: Retorna um dicionário onde cada chave é um ativo e cada valor é um dataframe das informações mais relevantes do ativo nos últimos 5 dias 
    """    
    dicionario_semana = {}
    for ativo in dic.keys():
        ticket=yf.Ticker(ativo)
        ticket_hist = ticket.history(period="5d")
        #chama a funçao conversão para corrigir o valor das colunas necessarias para BRL quando o ativo estiver cotado em outra moeda
        ticket_hist = conversao(ticket_hist)
        dicionario_semana[ativo] = ticket_hist
    return dicionario_semana


def cotacao_anual(dic):

    """A função recebe um dicionario que contém o nome das açoes como chaves e retorna a cotação do ativo no último ano
    Args:
        dic (Dictionary): Espera um dicionário onde as chaves são os códigos de cada ativo
    Returns:
        Dictionary: Retorna um dicionario onde cada chave é um ativo e cada valor é um dataframe das informações mais relevantes do ativo no último ano
    """  
    dicionario_anual={}
    for ativo in dic.keys():
        ticket=yf.Ticker(ativo)
        ticket_hist = ticket.history(period="1y")
        #chama a funçao conversão para corrigir o valor das colunas necessarias para BRL quando o ativo estiver cotado em outra moeda
        ticket_hist = conversao(ticket_hist)
        dicionario_anual[ativo] = ticket_hist
    return dicionario_anual

def conversao(ticket_hist):
    df = ticket_hist  
    df["Close"] = fator_conversao * df["Close"]
    df["Open"] = fator_conversao * df["Open"]
    return ticket_hist



def fator_conversao(dic):
    dicionario_moeda = {}
    for ativo in dic.keys():
        ticket=yf.Ticker(ativo)
        ticket_currency = ticket.info["currency"]
        if ticket_currency != "BRL":
            dicionario_moeda[ativo] = ticket_currency+"BRL=X"
        else:
            dicionario_moeda[ativo] = ticket_currency 
    for value in dicionario_moeda.values():
        fator = yf.Ticker(value)
        fator_info = fator.info["regularMarketPrice"] 
    return dicionario_moeda, fator_info


print(fator_conversao(dic))






"""
tickets=str()
for ativo in dic.keys():
    tickets = tickets+ativo+" "

ticket=yf.download(tickets, group_by="ticker", period="5d", progress=False)
for ativo in dic.keys():
    print(ticket[ativo].drop("Adj Close",axis=1).reset_index())
"""


#def conversão(f, dic):



#print(conversão(cotacao_semana(dic), dic))


#print(cotacao_semana(dic))


#yahooo query - info
#yf.download
#Tickers = list(dic.keys())
#a = "AMZN"
#df = yf.download(a, period="1mo", interval="1wk", progress=False)
#print(df)
#função dowload retorna dataframes diferentes pra 1 ativo ou pra mais ativos
#print(df.columns)

#acesso certo dataframe df=["close", "PETR4.SA"]



#currency = {"AMZN": "USD"}
#fator = currency.values()

#df["aa"] = fator[]"""


"""
tempo_inicial = time.time()
tempo_final = time.time()
print(f"{tempo_final - tempo_inicial} segundos")
calculo de tempo de execução
"""