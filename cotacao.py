import yfinance as yf

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
        dicionario_semana[ativo] = ticket_hist
    return dicionario_semana

def cotacao_anual(dic):
    """A função recebe um dicionario que contém o nome das açoes como chaves e retorna a cotação do ativo no último ano

    Args:
        dic (Dictionary): Espera um dicionário onde as chaves são os códigos de cada ativo

    Returns:
        list: Retorna uma lista onde cada elementos é um dataframe das informações mais relevantes do ativo no último ano
    """  
    lista_anual = []
    for x in dic.keys():
        y=yf.Ticker(x)
        y_hist = y.history(period="1y")
        lista_anual.append(y_hist)
    return lista_anual