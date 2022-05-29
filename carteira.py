from openpyxl import Workbook, load_workbook
from cotacao import cotacao_semana
from openpyxl.utils.dataframe import dataframe_to_rows
from criar_excel import criar_planilha

def criar_cabecalho(nome_arquivo):
    #Abre o arquivo
    planilha = load_workbook(nome_arquivo)
    #Seleciona a folha
    carteira = planilha["Carteira"]
    #Cria o título da carteira
    carteira["B3"]  = "Carteira de Ativos"
    carteira.merge_cells("B3:J7")
    #Criar espaço para QRCODE
    carteira["K3"] = "Valor Total"
    carteira.merge_cells("K3:L7")
    carteira["M3"] =  "QRCODE"
    carteira.merge_cells("M3:N7")
    #Define a linha final do cabeçalho
    carteira.cell(row=8,column=1, value=None)
    #Salva o arquivo
    planilha.save(nome_arquivo)

criar_planilha("Teste.xlsx")
criar_cabecalho("Teste.xlsx")