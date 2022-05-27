import requests
from bs4 import BeautifulSoup

URL = "https://zuilhose.github.io/A2-Intro-Comp/carteira1.html"
page = requests.get(URL)

soup = BeautifulSoup(page.content, "html.parser")

resultados_moedas = soup.find(class_="moedas")

elementos_moedas = resultados_moedas.find_all("td")

moedas = []
for elemento_moedas in elementos_moedas:
    moedas.append(elemento_moedas.text)

print(moedas)

resultados_acoes = soup.find(class_="acoes")

elementos_acoes = resultados_acoes.find_all("td")

acoes = []
for elemento_acoes in elementos_acoes:
    acoes.append(elemento_acoes.text)

print(acoes)


