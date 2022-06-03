from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, Series, Reference
from openpyxl.chart.marker import DataPoint
from datetime import datetime, timedelta
import yfinance as yf
import locale
import pandas
import calendar

locale.setlocale(locale.LC_ALL, ("pt_BR", "utf-8"))

def graf_barras(base, arquivo):
    
    #O gráfico de barras mostra quanto uma ação subiu ou desceu de valor no último ano, e para isso devemos primeiro adquirir os dados pelo yfinance, e em seguida montar o gráfico

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

        #É muito comum que em determinado dia, a bolsa não funcione por vários motivos, e por isso o Yfinance não terá dados referentes àquele dia específico, e nos entregará um dataframe vazio, por isso devemos tratar este caso. Caso o dataframe retornado seja vazio, iremos pegar os dados de um dia antes, repetidamente, até ser retornado um dataframe válido
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
    
    #Agora, vamos fazer a matemágica para chegar nos dados necessários
    dados = []
    for chave in valores_antigos.keys():
        x = valores_hoje[chave] / valores_antigos[chave]
        resultado = x - 1
        dados.append(resultado)
    print(dados)

    #E as categorias sendo as chaves:
    categorias = []
    for chave in valores_antigos.keys():
        categorias.append(chave)

    #Agora que temos os dados, entramos na etapa de montar o gráfico:

    #Carregamos a planilha e selecionamos a folha de trabalho
    planilha = load_workbook(arquivo)
    estatisticas = planilha["Estatísticas"]
    
    #Inserimos os dados na planilha
    linha_d = 2
    for dado in dados:
        linha_d += 1
        estatisticas[f"C{linha_d}"].value = dado        
    
    linha_c = 2
    for cat in categorias:
        linha_c += 1
        estatisticas[f"D{linha_c}"].value = cat

    #Selecionamos o tipo e os títulos do grafico
    grafic = BarChart()
    grafic.title = "Variação das Ações no Último Ano"
    grafic.type = "col"
    grafic.style = 9
    grafic.y_axis.title = "Variação"
    grafic.x_axis.title = "Ativos"
    grafic.x_axis.tickLblPos = "low"

    #Seleciona os dados do grafico
    dados_finais = Reference(estatisticas, min_col=3, min_row=3, max_col=3, max_row=f"{linha_d}")
    categorias_finais = Reference(estatisticas, min_col=4, min_row=3, max_col=4, max_row=f"{linha_c}")
    grafic.add_data(dados_finais, titles_from_data=False)
    grafic.set_categories(categorias_finais)

    #Modificando as cores das barras que representam valores_antigos menores que zero
    series = grafic.series[0]
    for i in range(len(dados)):
        if dados[i] < 0:
            pt = DataPoint(idx=i)
            pt.graphicalProperties.line.solidFill = 'DC143C'
            pt.graphicalProperties.solidFill = 'DC143C'
            series.dPt.append(pt)

    #Estilizando as séries
    grafic.legend = None

    #Adiciona o gráfico à planilha
    estatisticas.add_chart(grafic, "F3")
    planilha.save(arquivo)

"""==========================================================================================="""

# Dicionário de teste:
carteira = {"PETR4.SA":"10", "AMZN":"10", "AAPL":"100", "KO":"100"}

graf_barras(carteira, "teste.xlsx")
