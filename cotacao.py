import yfinance as yf
import time

dic = {"AMZN":"10", "KO":"10", "AAPL":"10", "PETR4.SA":"12"}

def buscar_fator(dic):
    dicionario_fatores = {}
    for ativo in dic.keys():
        ticket=yf.Ticker(ativo)
        ticket_currency = ticket.info["currency"]
        if ticket_currency != "BRL":
            ticket_currency = ticket_currency+"BRL=X"
            fator = yf.Ticker(ticket_currency)
            fator_info = fator.info["regularMarketPrice"]
            dicionario_fatores[ativo] = fator_info
        else:
            dicionario_fatores[ativo] = 1
    return print(dicionario_fatores)


def conversao(ticket_hist, dic):
    #chama a função conversão para receber um dicionario com os fatores de correção de cada ativo 
    fatores = buscar_fator(dic)
    for ativo in dic.keys():
        fator_conversao = fatores[ativo]
        #evita a execução para quando o ativo ja esta cotado em BRL para otimização do código
        if fator_conversao != 1:
            df = ticket_hist  
            df["Open"] = fator_conversao * df["Open"]
            df["High"] = fator_conversao * df["High"]
            df["Low"] = fator_conversao * df["Low"]
            df["Close"] = fator_conversao * df["Close"]
    return ticket_hist

def cotacao_semana(dic):
    tempo1 = time.time()
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
        ticket_hist = conversao(ticket_hist, dic)
        dicionario_semana[ativo] = ticket_hist
    tempo2 = time.time()
    print(f"{tempo2-tempo1} segundos")
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
        ticket_hist = conversao(ticket_hist, dic)
        dicionario_anual[ativo] = ticket_hist
    return dicionario_anual




