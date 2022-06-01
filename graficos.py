from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, Series, Reference
from openpyxl.chart.marker import DataPoint
from datetime import datetime, timedelta
import yfinance as yf
import locale
locale.setlocale(locale.LC_ALL, ("pt_BR", "utf-8"))
import pandas

# Dicionário de teste:
carteira = {"PETR4.SA":"10", "AMZN":"10", "AAPL":"100", "KO":"100"}

def graf_barras(base, planilha):
 
    #Código para pegar a data de um ano atrás 
    x = 366  
    #Esse x sera parâmetro de debug mais pra frente
    data_inicio = datetime.now() - timedelta(days=x)
    data_final = data_inicio + timedelta(days=1)
    #Também devemos formatar a string de forma que o Yfinance entenda
    data_inicio = data_inicio.strftime("%Y-%m-%d")
    data_final = data_final.strftime("%Y-%m-%d")
    
    valores = {}
    for ativo in base.keys():
        ticket = yf.Ticker(ativo)
        ticket_hist = ticket.history(start=data_inicio, end=data_final)

        while ticket_hist.size == 0:
            x += 1
            inicio_altern = datetime.now() - timedelta(days=x)
            final_altern = inicio_altern + timedelta(days=1)
            ticket_hist = ticket.history(start=inicio_altern, end=final_altern)

        x = 366
        valores[ativo] = ticket_hist.Close[0]

    print(valores)
    
    
"""==========================================================================================="""

graf_barras(carteira, "aaa")
