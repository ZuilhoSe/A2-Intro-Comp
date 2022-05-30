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
    carteira.merge_cells("B3:K7")
    #Criar espaço para QRCODE
    carteira["L3"] = "Valor Total"
    carteira.merge_cells("L3:M7")
    carteira["N3"] =  "QRCODE"
    carteira.merge_cells("N3:O7")
    #Salva o arquivo
    planilha.save(nome_arquivo)

def criar_tabela_ativo(nome_arquivo,ativo,data_frame,linha_inicial):
    #Abre o arquivo
    planilha = load_workbook(nome_arquivo)
    #Seleciona a folha
    carteira = planilha["Carteira"]
    #Cria o título do ativo
    linha_titulo = linha_inicial + 1
    carteira.cell(row = linha_titulo,column = 3, value = ativo)
    carteira.merge_cells(start_row = linha_titulo, start_column = 3, end_row = linha_titulo+1, end_column = 10)
    #Define a linha inicial da tabela
    linha_tabela = linha_inicial + 3 
    carteira.cell(row = linha_tabela,column = 3, value = None)
    #Acrescenta a tabela dataframe
    for linha in dataframe_to_rows(data_frame,index=True,header=True):
        carteira.append(linha)
    #Salva o arquivo
    planilha.save(nome_arquivo)

def criar_corpo_carteira(nome_arquivo,dicionario_ativos):
    cotacao = cotacao_semana(dicionario_ativos)
    linha_inicial = 9
    for ativo,data_frame in cotacao.items():
        criar_tabela_ativo(nome_arquivo,ativo,data_frame,linha_inicial)
        linha_inicial += 13

nome_arquivo = "Teste.xlsx"
criar_planilha(nome_arquivo)
criar_cabecalho(nome_arquivo)
dicionario_teste = {"MGLU3.SA":"1000000", "KO":"98987"}
criar_corpo_carteira(nome_arquivo,dicionario_teste)
