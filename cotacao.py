import yfinance as yf
import time

dic = {"AMZN":"100"}
fator_conversao = 2


def buscar_fator(dic):
    dicionario_moeda = {}
    dicionario_fatores = {}
    for ativo in dic.keys():
        ticket=yf.Ticker(ativo)
        ticket_currency = ticket.info["currency"]
        if ticket_currency != "BRL":
            dicionario_moeda[ativo] = ticket_currency+"BRL=X"
        else:
            dicionario_moeda[ativo] = ticket_currency 
        for value in dicionario_moeda.values():
            if value != "BRL":
                fator = yf.Ticker(value)
                fator_info = fator.info["regularMarketPrice"]
                dicionario_fatores[ativo] = fator_info
            else:
                dicionario_fatores[ativo] = 1
    return dicionario_fatores



def conversao(ticket_hist, dic):
    fatores = buscar_fator(dic)
    for ativo in dic.keys():
        fator_conversao = fatores[ativo]
        df = ticket_hist  
        df["Close"] = fator_conversao * df["Close"]
        df["Open"] = fator_conversao * df["Open"]
    return ticket_hist

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
        ticket_hist = conversao(ticket_hist, dic)
        dicionario_semana[ativo] = ticket_hist
    return dicionario_semana

print(cotacao_semana(dic))

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




