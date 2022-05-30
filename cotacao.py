import yfinance as yf
import time

dic = {"PETR4.SA":"10", "AMZN":"10"}

def buscar_fator(ativo):
    ticket=yf.Ticker(ativo)
    ticket_currency = ticket.info["currency"]
    if ticket_currency != "BRL":
        ticket_currency = ticket_currency+"BRL=X"
        fator = yf.Ticker(ticket_currency)
        fator_info = fator.info["regularMarketPrice"]
        return fator_info
    else:
        return 1



def conversao(ticket_hist, fator):
    if fator != 1:
        ticket_hist  
        ticket_hist ["Open"] = fator * ticket_hist ["Open"]
        ticket_hist ["High"] = fator * ticket_hist ["High"]
        ticket_hist ["Low"] = fator * ticket_hist ["Low"]
        ticket_hist ["Close"] = fator * ticket_hist ["Close"]
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
        fator = buscar_fator(ativo)
        ticket_hist = conversao(ticket_hist, fator)

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
        ticket_hist = conversao(ticket_hist, dic)
        dicionario_anual[ativo] = ticket_hist
    return dicionario_anual




