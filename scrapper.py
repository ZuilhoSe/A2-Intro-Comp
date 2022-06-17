import requests
from bs4 import BeautifulSoup

#Função que retorna o HTML parseado da página requerida;
def ler_carteira(URL):
    """Função que retorna o objeto Soup;

    Args:
        URL (string): Url do site com a carteira de investimentos;

    Returns:
        soup: Objeto do tipo soup, parseado;
    """    
    page = requests.get(URL)
    soup = BeautifulSoup(page.content, "html.parser")
    return soup

#Função que retorna uma lista com as moedas existentes na carteira e a suas quantidades;
def buscar_moedas(soup):
    """Função que busca os ativos do tipo moeda na carteira;

    Args:
        soup (soup): Objeto do tipo soup, parseado;

    Returns:
        list: Lista com as moedas e suas quantias;
    """    
    resultados_moedas = soup.find(class_="moeda")
    
    elementos_moedas = resultados_moedas.find_all("td")

    moedas = []
    for elemento_moedas in elementos_moedas:
        moedas.append(elemento_moedas.text)

    return moedas

#Função que retorna uma lista com as ações existentes na carteira e a suas quantidades;
def buscar_acoes(soup):
    """Função que busca os ativos do tipo ações na carteira;

    Args:
        soup (soup): Objeto do tipo soup, parseado;

    Returns:
        list: Lista com as ações e suas quantias;
    """    
    resultados_acoes = soup.find(class_="acao")
    
    elementos_acoes = resultados_acoes.find_all("td")

    acoes = []
    for elemento_acoes in elementos_acoes:
        acoes.append(elemento_acoes.text)

    return acoes

#Função que junta as listas de moedas e ações;
def juntar_moedas_acoes(moedas, acoes):
    """Função que retorna a lista conjunta de moedas e ações;

    Args:
        moedas (list): Lista com as moedas e suas quantias;
        acoes (list): Lista com as ações e suas quantias;

    Returns:
        list: Lista conjunta com as ações, as moedas e suas quantias;
    """    
    lista = moedas + acoes
    return lista    

#Função que transforma a lista de ações e moedas em um dicionário;
def saida(lista):
    """Função de saída final. Transforma a lista com moedas e ações em um dicionário onde a chave é o ativo e valor é a quantidade;

    Args:
        lista (list): Lista conjunta com as ações, as moedas e suas quantias;

    Returns:
        dict: Dicionário onde a chave é o nome do ativo e o valor é a sua quantidade;
    """    
    carteira[lista[0]] = lista[1]

    if len(lista) == 2:
        return carteira
    else:
        return saida(lista[2:])

carteira = {}

"""
Aqui temos um teste básico para vermos se as funções estão funcionando:
soup = ler_carteira("https://zuilhose.github.io/A2-Intro-Comp/")
acoes = buscar_acoes(soup)
moedas = buscar_moedas(soup)
lista = juntar_moedas_acoes(acoes, moedas)
carteira = saida(lista)
print(carteira)
"""