from openpyxl import load_workbook
from openpyxl.chart import BarChart, Series, Reference
from openpyxl.chart.marker import DataPoint
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.axis import TextAxis, ChartLines
from openpyxl.chart.chartspace import ChartSpace, ChartContainer
from openpyxl.chart.plotarea import PlotArea
from openpyxl.drawing.line import LineProperties
from openpyxl.chart.layout import Layout, ManualLayout
from datetime import datetime, timedelta
import yfinance as yf
import locale
import calendar
from carteira import aplicar_estilo_area
from openpyxl.styles import NamedStyle, PatternFill, Font

from cotacao import cotacao_anual

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
        ticket_hist = ticket.history(start=data_inicio, end=data_final, debug = False)
        ticket_hist.fillna(-1, inplace = True)

        # É muito comum que em determinado dia, a bolsa não funcione por vários motivos, e por isso o Yfinance 
        # não terá dados referentes àquele dia específico, e nos entregará um dataframe vazio, por isso 
        # devemos tratar este caso. Caso o dataframe retornado seja vazio, iremos pegar os dados de um dia 
        # antes, repetidamente, até ser retornado um dataframe válido.

        # Também foi necessário tratar um caso em que o dataframe é retornado com alguns valores faltantes,
        # com as células contendo NaN. Para isso foi usado o método fillna do dataframe, e um Or no while.

        while ticket_hist.size == 0 or ticket_hist.Close[0] == -1:
            dias += 1
            inicio_altern = datetime.now() - timedelta(days=dias)
            final_altern = inicio_altern + timedelta(days=1)
            ticket_hist = ticket.history(start=inicio_altern, end=final_altern, debug = False)
            ticket_hist.fillna(-1, inplace = True)

        valores_antigos[ativo] = ticket_hist.Close[0]

    print(valores_antigos)
    
    #Agora, fazemos o mesmo processo para os dados de cada ativo no dia atual
    valores_hoje = {}
    for ativo in base.keys():
        dia_hoje = 1
        ticket = yf.Ticker(ativo)
        ticket_hist = ticket.history(period=f"{dia_hoje}d", debug = False)
        ticket_hist.fillna(-1, inplace = True)

        while ticket_hist.size == 0 or ticket_hist.Close[0] == -1 :
            dia_hoje += 1
            ticket_hist = ticket.history(period=f"{dia_hoje}d", debug = False)
            ticket_hist.fillna(-1, inplace = True)

        valores_hoje[ativo] = ticket_hist.Close[0]

    print(valores_hoje)
    
    #Agora, vamos fazer a matemágica para chegar nos dados necessários
    dados = []
    for chave in valores_antigos.keys():
        resultado = (valores_hoje[chave] / valores_antigos[chave]) - 1
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
    bar_grafic = BarChart()
    bar_grafic.title = "Variação das Ações no Último Ano"
    bar_grafic.type = "col"
    bar_grafic.y_axis.title = "Variação"
    bar_grafic.y_axis.number_format = "0%"
    bar_grafic.x_axis.title = "Ativos"
    bar_grafic.x_axis.tickLblPos = "low"

    #Seleciona os dados do grafico
    dados_finais = Reference(estatisticas, min_col=3, min_row=3, max_col=3, max_row=f"{linha_d}")
    categorias_finais = Reference(estatisticas, min_col=4, min_row=3, max_col=4, max_row=f"{linha_c}")
    bar_grafic.add_data(dados_finais, titles_from_data=False)
    bar_grafic.set_categories(categorias_finais)

    #Modificando as cores das barras que representam valores menores que zero (Isso cria uma série nova pra cada barra negativa, então tecnicamente é uma gambiarra)
    series = bar_grafic.series[0]
    for i in range(len(dados)):
        if dados[i] < 0:
            pt = DataPoint(idx=i)
            pt.graphicalProperties.line.solidFill = "DC143C"
            pt.graphicalProperties.solidFill = "DC143C"
            series.dPt.append(pt)
        else:
            pt = DataPoint(idx=i)
            pt.graphicalProperties.line.solidFill = "00A300"
            pt.graphicalProperties.solidFill = "00A300"
            series.dPt.append(pt)
            
    #Estilizando
    bar_grafic.legend = None
    bar_grafic.style = 9
    # bar_grafic.plot_area.graphicalProperties = GraphicalProperties(solidFill="36454F") 
    bar_grafic.x_axis.graphicalProperties = GraphicalProperties(ln=LineProperties(w=3))
    bar_grafic.y_axis.majorGridlines.graphicalProperties = GraphicalProperties(ln=LineProperties(prstDash="dash"))
    
    # estilo_esconder = NamedStyle(name = "estilo_esconder")
    # estilo_esconder.fill = PatternFill(fill_type = "solid", fgColor="FFFFFF")
    # estilo_esconder.font = Font(color="FFFFFF")
    aplicar_estilo_area(estatisticas, 3, linha_d, 3, 4, "estilo_esconder")

    #Adiciona o gráfico à planilha
    estatisticas.add_chart(bar_grafic, "F3")
    planilha.save(arquivo)
    
    print("Cabou")


def graf_stock(base, arquivo):

    #Essa função irá criar um Stock Chart com as informações de cada ativo no último ano, de semana a semana

    #Para teste vou começar importando a função do módulo cotacao. Depois isso poderá ser descartado.
    
    dict = cotacao_anual(base)



"""==========================================================================================="""


# Dicionário de teste:
carteira = {"PETR4.SA":"10", "AMZN":"10", "AAPL":"100", "KO":"100"}

# graf_barras(carteira, "teste.xlsx")

# graf_stock(carteira, "teste2.xlsx")
