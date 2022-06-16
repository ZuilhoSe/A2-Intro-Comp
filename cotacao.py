import yfinance as yf
import time

def buscar_fator(ticket, ticket_hist, periodo, intervalo="1d"):
    """A função recebe uma string que é o código da ação no yahoo finance encontra a currency desse ativo e caso ele seja diferente de BRL executa uma concatenação e busca uma informação sobre o ativo

    Args:
        ativo (str): Espera uma string que seja um código de ativo valido no yahoo finance 

    Returns:
        float: Retorna um float que é o valor da moeda do ativo em relação ao real
    """    
    #encontra a moeda em que o ativos esta cotado
    ticket_currency = ticket.info["currency"]
    #caso não esteja cotado em BRL o bloco abaixo concatena a string da moeda em que o ativo esta cotado com "BRL=X"
    if ticket_currency != "BRL":
        ticket_currency = ticket_currency+"BRL=X"
        #busca o ticket usando a string concatenada
        fator = yf.Ticker(ticket_currency)
        #encontra o valor da moeda em relaçao ao real 
        fator_info = fator.history(period=periodo, interval=intervalo)
        return fator_info
    else:
        return ticket_hist

def conversao(ticket_hist, fator):
    """converte as colunas desejas do data frame para apresentar o valor em real

    Args:
        ticket_hist (pandas.core.frame.DataFrame): recebe um dataframe de pandas que tera as colunas desejadas multiplicadas pela cotação da moeda do ativo em relação ao real
        fator (float): é a cotação da moeda do ativo em relação ao real

    Returns:
        pandas.core.frame.DataFrame: retorna o dataframe com as colunas alteradas
    """    
    #se o fator de correção for diferente de um aplica esse fator nas colunas desejadas do dataframe se for igual a 1 apenas retorna o primeiro argumento
    ticket_hist ["Open"] = fator["Open"] * ticket_hist ["Open"]
    ticket_hist ["High"] = fator["High"] * ticket_hist ["High"]
    ticket_hist ["Low"] = fator["Low"] * ticket_hist ["Low"]
    ticket_hist ["Close"] = fator["Close"] * ticket_hist ["Close"]
    return ticket_hist


def cotacao_semana(dic):
    #clock=time.time()
    """A função recebe um dicionario que contém o nome das açoes como chaves e retorna a cotação do ativo na última semana útil
    Args:
        dic (Dictionary): Espera um dicionário onde as chaves são os códigos de cada ativo
    Returns:
        Dictionary: Retorna um dicionário onde cada chave é um ativo e cada valor é um dataframe das informações mais relevantes do ativo nos últimos 5 dias 
    """    
    dicionario_semana = {}
    periodo="5d"
    for ativo in dic.keys():
        ticket=yf.Ticker(ativo)
        ticket_hist = ticket.history(period="5d")
        #chama a funçao buscar_fator para encontrar o valor em relação ao real da moeda em que o ativo está cotado
        fator = buscar_fator(ticket, ticket_hist, periodo)
        #chama a função conversao que usa o resultado de buscar_fator para converter as colunas necessarias para real quando o ativo esta cotado em outra moeda
        ticket_hist = conversao(ticket_hist, fator)

        dicionario_semana[ativo] = ticket_hist
    #clock2=time.time()
    #print(f"{clock2-clock} segundos de execução")
    return dicionario_semana

def cotacao_anual(dic):
    """A função recebe um dicionario que contém o nome das açoes como chaves e retorna a cotação do ativo no último ano
    Args:
        dic (Dictionary): Espera um dicionário onde as chaves são os códigos de cada ativo
    Returns:
        Dictionary: Retorna um dicionario onde cada chave é um ativo e cada valor é um dataframe das informações mais relevantes do ativo no último ano
    """  
    dicionario_anual={}
    periodo="1y"
    intervalo="1wk"
    for ativo in dic.keys():
        ticket=yf.Ticker(ativo)
        ticket_hist = ticket.history(period=periodo, interval=intervalo)
        #chama a funçao buscar_fator para encontrar o valor em relação ao real da moeda em que o ativo está cotado
        fator = buscar_fator(ticket, ticket_hist, periodo, intervalo)
        #chama a funçao conversao para corrigir o valor das colunas necessarias para real quando o ativo estiver cotado em outra moeda
        ticket_hist = conversao(ticket_hist, fator)
        dicionario_anual[ativo] = ticket_hist
    return dicionario_anual


def cotacao_atual(dic):
    """A função recebe um dicionario que contém o nome das açoes como chaves e retorna a cotação atual do ativo

    Args:
        dic (Dictionary): Espera um dicionário onde as chaves são os códigos de cada ativo

    Returns:
        Dictionary: Retorna um dicionário onde cada chave é o código de um ativo e os valores são o valor atual de cada ativo convertido para real
    """    
    dicionario_atual={}
    for ativo in dic.keys():
        ticket=yf.Ticker(ativo)
        ticket_info=ticket.info["regularMarketPrice"]
        #encontra o valor da moeda em que o ativo está cotado em relação ao real
        fator = buscar_fator(ticket)
        #converte o valor do ativo para real 
        ticket_info = ticket_info*fator
        dicionario_atual[ativo] = ticket_info
    return dicionario_atual

# use o dicionario a baixo como teste 
#dic = {"PETR4.SA":"10", "AMZN":"10"}


