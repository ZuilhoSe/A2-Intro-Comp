from ast import Try
from types import NoneType
from pandas import DataFrame
import yfinance as yf

def buscar_fator(ticket, ticket_hist, periodo, intervalo="1d"):
    """Busca um dataframe contendo os dados historicos dos valores da moeda na qual o ativo é cotado em relação ao real e retorna esse data frame para que o data frame do ativo seja corrigido para real.

    Args:
        ticket (yfinance.ticker.Ticker):Recebe um Ticker do yfinance que sera usado para consultar as informações necessaria para a correção de valor dos ativos.
        ticket_hist (pandas.core.frame.DataFrame): Recebe o Data frame do ativo e retorna ele caso seja um ativo cotado em BRL.
        periodo (str): O periodo desejado.
        intervalo (str, optional): O intervalo desejado. Defaults to "1d".

    Returns:
        pandas.core.frame.DataFrame: Retorna um Data Frame com os valores históricos da moeda em que o ativo esta cotado em relação ao real.
        pandas.core.frame.DataFrame: Retorna o Data frame com os valores historicos do ATIVO.
    """ 
    #encontra a moeda em que o ativos esta cotado
    ticket_currency = ticket.info["currency"]
    #caso não esteja cotado em BRL o bloco abaixo concatena a string da moeda em que o ativo esta cotado com "BRL=X"
    if type(ticket_currency)==NoneType:
        ticket_currency= "BRL"
    if ticket_currency != "BRL":
        ticket_currency = ticket_currency+"BRL=X"
        #busca o ticket usando a string concatenada
        fator = yf.Ticker(ticket_currency)
        #encontra a cotação histórica da moeda em relação ao real 
        fator_info = fator.history(period=periodo, interval=intervalo)
        return fator_info
    #caso a moeda esteja cotada em BRL nenhuma ação é tomada e o data frame do ativo é retornado
    else:
        return ticket_hist

def buscar_fator_atual(ticket):
    """Semelhante a função buscar_fator encontra um fator de correção para o valor atual de um ativo.

    Args:
        ticket (yfinance.ticker.Ticker): Recebe um Ticker do yfinance que sera usado para consultar as informações necessaria para a correção de valor dos ativos.

    Returns:
        float: Retorna um float que será o fator de conversão do valor do ativo para BRL.
    """    
    ticket_currency = ticket.info["currency"]
    if type(ticket_currency)==NoneType:
        ticket_currency= "BRL"
    if ticket_currency != "BRL":
        ticket_currency = ticket_currency+"BRL=X"
        #busca o ticket usando a string concatenada
        fator = yf.Ticker(ticket_currency)
        #encontra o valor da moeda em relaçao ao real 
        fator_info = fator.info["regularMarketPrice"]
        return fator_info
    else:
        return 1


def conversao(ticket_hist, fator):
    """Converte as colunas desejas do data frame para apresentar o valor em real.

    Args:
        ticket_hist (pandas.core.frame.DataFrame): Recebe um dataframe pandas que contem os dados historicos que serão usados para conversão.
        fator (pandas.core.frame.DataFrame): é a cotação da moeda do ativo em relação ao real.

    Returns:
        pandas.core.frame.DataFrame: retorna o dataframe com as colunas alteradas.
    """
    igualdade=DataFrame.equals(fator,ticket_hist)
    if igualdade == True :
        return ticket_hist
    else:
        #se o fator de correção não for o proprio data frame corrige os valores das colunas abaixo para real
        ticket_hist ["Open"] = fator["Open"] * ticket_hist ["Open"]
        ticket_hist ["High"] = fator["High"] * ticket_hist ["High"]
        ticket_hist ["Low"] = fator["Low"] * ticket_hist ["Low"]
        ticket_hist ["Close"] = fator["Close"] * ticket_hist ["Close"]
    return ticket_hist


def cotacao_semanal(dic):
    """A função recebe um dicionario que contém o nome das açoes como chaves e retorna a cotação do ativo na última semana útil.
    Args:
        dic (Dictionary): Espera um dicionário onde as chaves são os códigos de cada ativo.
    Returns:
        Dictionary: Retorna um dicionário onde cada chave é um ativo e cada valor é um dataframe das informações mais relevantes do ativo nos últimos 5 dias.
    """    
    dicionario_semanal = {}
    periodo="5d"
    for ativo in dic.keys():
        ticket=yf.Ticker(ativo)
        ticket_hist = ticket.history(period=periodo)
        #caso um ativo não esteja listado no Yahoo finance ticket_hist será um Data frame vazio, nesses casos o for deve ignorar essa entrada e seguir para os proximos ativos 
        if not ticket_hist.empty:
            #chama a funçao buscar_fator para encontrar o valor em relação ao real da moeda em que o ativo está cotado
            fator = buscar_fator(ticket, ticket_hist, periodo)
            #chama a função conversao que usa o resultado de buscar_fator para converter as colunas necessarias para real quando o ativo esta cotado em outra moeda
            ticket_hist = conversao(ticket_hist, fator)
        else:
            continue
        dicionario_semanal[ativo] = ticket_hist
    return dicionario_semanal

def cotacao_anual(dic):
    """A função recebe um dicionario que contém o nome das açoes como chaves e retorna a cotação do ativo no último ano.
    Args:
        dic (Dictionary): Espera um dicionário onde as chaves são os códigos de cada ativo.
    Returns:
        Dictionary: Retorna um dicionario onde cada chave é um ativo e cada valor é um dataframe das informações mais relevantes do ativo no último ano.
    """  
    dicionario_anual={}
    periodo="1y"
    intervalo="1wk"
    for ativo in dic.keys():
        ticket=yf.Ticker(ativo)
        ticket_hist = ticket.history(period=periodo, interval=intervalo)
        #caso um ativo não esteja listado no Yahoo finance ticket_hist será um Data frame vazio, nesses casos o for deve ignorar essa entrada e seguir para os proximos ativos
        if not ticket_hist.empty:
            #chama a funçao buscar_fator para encontrar o valor em relação ao real da moeda em que o ativo está cotado
            fator = buscar_fator(ticket, ticket_hist, periodo, intervalo)
            #chama a função conversao que usa o resultado de buscar_fator para converter as colunas necessarias para real quando o ativo esta cotado em outra moeda
            ticket_hist = conversao(ticket_hist, fator)
        else:
            continue
        dicionario_anual[ativo] = ticket_hist
    return dicionario_anual


def cotacao_atual(dic):
    """A função recebe um dicionario que contém o nome das açoes como chaves e retorna a cotação atual do ativo.

    Args:
        dic (Dictionary): Espera um dicionário onde as chaves são os códigos de cada ativo.

    Returns:
        Dictionary: Retorna um dicionário onde cada chave é o código de um ativo e os valores são o valor atual de cada ativo convertido para real.
    """    
    dicionario_atual={}
    for ativo in dic.keys():
        ticket=yf.Ticker(ativo)
        #Caso algum ativo não esteja listado no yahoo finance será retornado KeyError nesses casos queremos que nosso for ignore esse ativo e continue para os demais
        try:
            ticket_info=ticket.info["regularMarketPrice"]
            #encontra o valor da moeda em que o ativo está cotado em relação ao real
            fator = buscar_fator_atual(ticket)
            #converte o valor do ativo para real 
            ticket_info = ticket_info*fator
            dicionario_atual[ativo] = ticket_info
        except KeyError:
            continue
    return dicionario_atual

# use o dicionario a baixo como teste 
#dic = {'AVST.L': 'GBp', 'ANTO.L': 'GBp', 'PEsTZ3.SA': 'GBp', '9988d.HK': 'GBp', '07030.HK': 'GBp', 'ARS4BRL=X': 'GBp', 'INR1BRL=X': 'GBp', 'CHFB3RL=X': 'GBp', 'TWDBRL=X': None}
