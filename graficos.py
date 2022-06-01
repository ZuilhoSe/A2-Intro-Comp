from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, Series, Reference
from openpyxl.chart.marker import DataPoint
from datetime import datetime, timedelta
import yfinance as yf
import locale
locale.setlocale(locale.LC_ALL, ("pt_BR", "utf-8"))
import pandas
import calendar

# Dicionário de teste:
carteira = {"PETR4.SA":"10", "AMZN":"10", "AAPL":"100", "KO":"100"}

def graf_barras(base, planilha):
    
    #O gráfico de barras mostra quanto uma ação subiu ou desceu de valor no último ano, e para isso devemos primeiro adquirir os dados pelo dia_hoje finance, e em seguida montar o gráfico

    #Primeiro, o código para pegar a data de um ano atrás:
    dias = 365
    data_inicio = datetime.now() - timedelta(days=dias)
    
    #Vamos tratar o caso de termos um ano bissexto, por que não?
    ano = int(data_inicio.strftime("%Y"))
    if calendar.isleap(ano) == True:
        dias += 1
        data_inicio = datetime.now() - timedelta(days=dias)
    
    data_final = data_inicio + timedelta(days=1)

    #Também devemos formatar a string na forma que o Yfinance entende
    data_inicio = data_inicio.strftime("%Y-%m-%d")
    data_final = data_final.strftime("%Y-%m-%d")
    
    #Agora, vamos pegar os valores de cada ativo um ano atrás
    valores_antigos = {}
    for ativo in base.keys():
        ticket = yf.Ticker(ativo)
        ticket_hist = ticket.history(start=data_inicio, end=data_final)

        #É muito comum que em determinado dia, a bolsa não funcione por vários motivos, e por isso o Yfinance não terá dados referentes àquele dia específico, e nos entregará um dataframe vazio, por isso devemos tratar este caso. Decidi tratar de tal forma que, caso o dataframe retornado seja vazio, iremos pegar os dados de um dia antes, repetidamente, até ser retornado um dataframe válido
        while ticket_hist.size == 0:
            dias += 1
            inicio_altern = datetime.now() - timedelta(days=dias)
            final_altern = inicio_altern + timedelta(days=1)
            ticket_hist = ticket.history(start=inicio_altern, end=final_altern)

        valores_antigos[ativo] = ticket_hist.Close[0]

    print(valores_antigos)
    
    #Agora, fazemos o mesmo processo para os dados de cada ativo no dia atual
    valores_hoje = {}
    for ativo in base.keys():
        dia_hoje = 1
        ticket = yf.Ticker(ativo)
        ticket_hist = ticket.history(period=f"{dia_hoje}d")

        while ticket_hist.size == 0:
            dia_hoje += 1
            ticket_hist = ticket.history(period=f"{dia_hoje}d")

        valores_hoje[ativo] = ticket_hist.Close[0]

    print(valores_hoje)
    
    """
    #Carregar a planilha e selecionar a folha de trabalho
    planilha = load_workbook(planilha)
    folha = planilha.active
    
    
    #Seleciona o tipo e os titulos do grafico
    grafic = BarChart()
    grafic.title = "Variação das Ações no Último Ano"
    grafic.type = "col"
    grafic.style = 7
    grafic.y_axis.title = "Variação"
    grafic.x_axis.title = "Ativo"
    grafic.x_axis.tickLblPos = "low"

    #Seleciona os dados do grafico
    dados = Reference(folha, min_col=3, min_row=2, max_col=3, max_row=14)
    nomes = Reference(folha, min_col=2, min_row=3, max_col=2, max_row=14)
    grafic.add_data(dados, titles_from_data=True)
    grafic.set_categories(nomes)

    #Modificando as cores das barras que representam valores_antigos menores que zero
    rows = []
    for row in folha.iter_rows(min_row=3, min_col=2, max_col=3, max_row=14, values_only=True):
        rows.append(row)

    series = grafic.series[0]
    for i in range(len(rows)):
        if rows[i][1] < 0:
            pt = DataPoint(idx=i)
            pt.graphicalProperties.line.solidFill = 'DC143C'
            pt.graphicalProperties.solidFill = 'DC143C'
            series.dPt.append(pt)

    #Estilizando as séries
    grafic.legend = None
    grafic.shape = 1

    #Adiciona o gráfico à planilha
    folha.add_chart(grafic, "B20")

    planilha.save("teste 2.xlsx")

    """
    
"""==========================================================================================="""

graf_barras(carteira, "aaa")
