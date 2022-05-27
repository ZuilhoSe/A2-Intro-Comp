import requests
from bs4 import BeautifulSoup

#Função que retorna o HTML parseado da página requerida;
def ler_carteira(URL):
    page = requests.get(URL)
    soup = BeautifulSoup(page.content, "html.parser")
    return soup

#Função que retorna uma lista com as moedas existentes na carteira e a suas quantidades;
def buscar_moedas(soup):
    resultados_moedas = soup.find(class_="moeda")

    elementos_moedas = resultados_moedas.find_all("td")

    moedas = []
    for elemento_moedas in elementos_moedas:
        moedas.append(elemento_moedas.text)

    return moedas

#Função que retorna uma lista com as ações existentes na carteira e a suas quantidades;
def buscar_acoes(soup):
    resultados_acoes = soup.find(class_="acao")

    elementos_acoes = resultados_acoes.find_all("td")

    acoes = []
    for elemento_acoes in elementos_acoes:
        acoes.append(elemento_acoes.text)

    return acoes

#Função que junta as listas de moedas e ações;
def juntar_moedas_acoes(moedas, acoes):
    lista = moedas + acoes
    return lista    

#Função que transforma a lista de ações e moedas em um dicionário;
def saida(lista):
    carteira[lista[0]] = lista[1]

    if len(lista) == 2:
        return carteira
    else:
        return saida(lista[2:])

carteira = {}

"""
Aqui temos um teste básico para vermos se as funções estão funcionando:

soup = ler_carteira("https://zuilhose.github.io/A2-Intro-Comp/carteira3.html")
acoes = buscar_acoes(soup)
moedas = buscar_moedas(soup)
lista = juntar_moedas_acoes(acoes, moedas)
carteira = saida(lista)
print(carteira)
"""