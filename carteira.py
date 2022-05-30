from json import load
from openpyxl import Workbook, load_workbook
from cotacao import cotacao_semana
from openpyxl.utils.dataframe import dataframe_to_rows
from criar_excel import criar_planilha

def criar_cabecalho(nome_folha):
    """Cria o cabeçalho da planilha

    Args:
        nome_folha (openpyxl.worksheet.worksheet.Worksheet): deve estar no formato load_workbook(nome_arquivo)["nome_folha"]
    """    
    #Cria o título da carteira
    nome_folha["B3"]  = "Carteira de Ativos"
    nome_folha.merge_cells("B3:K7")
    #Criar espaço para QRCODE
    nome_folha["L3"] = "Valor Total"
    nome_folha.merge_cells("L3:M7")
    nome_folha["N3"] =  "QRCODE"
    nome_folha.merge_cells("N3:O7")

def criar_tabela_ativo(nome_folha,ativo,data_frame,linha_inicial,coluna_inicial):
    """Cria na planilha a tabela por ativo

    Args:
        nome_folha (openpyxl.worksheet.worksheet.Worksheet): deve estar no formato load_workbook(nome_arquivo)["nome_folha"]
        ativo (string): nome do ativo que será o título
        data_frame (dataframe): dataframe que se tornará a tabela
        linha_inicial (int): linha em que começa o bloco da tabela
        coluna_inicial (int): coluna em que começa o bloco
    """    
    #Cria o título do ativo
    linha_titulo = linha_inicial + 1
    nome_folha.cell(row = linha_titulo,column = coluna_inicial, value = ativo)
    nome_folha.merge_cells(start_row = linha_titulo, start_column = coluna_inicial, end_row = linha_titulo+1, end_column = coluna_inicial +  7)
    #Define a linha inicial da tabela
    linha_tabela = linha_inicial + 4 
    nome_folha.cell(row = linha_tabela,column = coluna_inicial, value = None)
    #Transforma o dataframe em uma lista de linhas
    linhas_data_frame = list(dataframe_to_rows(data_frame,index=True,header=True))
    """
    O rótulo date estava uma linha abaixo como frozen list, foi necessário converte-lo e adicioná-lo a linha correta
    """
    lista_rotulo_date = list(linhas_data_frame[1])
    rotulo_date =  str(lista_rotulo_date[0])
    linhas_data_frame[0][0] = rotulo_date
    del linhas_data_frame[1]
    #Acrescenta a tabela dataframe
    for linha in linhas_data_frame:
        #Pega cada lista e coloca em uma linha, com cada item em uma coluna
        coluna_item = coluna_inicial
        for item_coluna in linha:
            nome_folha.cell(row = linha_tabela,column = coluna_item, value = item_coluna)
            coluna_item += 1
        linha_tabela += 1    

def criar_blocos(linha_inicial,coluna_inicial,titulo,valor,nome_folha):
    """Cria um dois blocos no excel, com tammanho 2x2 um ao lado do outro

    Args:
        linha_inicial (int): linha em que começa os blocos da tabela
        coluna_inicial (int): coluna em que começa o bloco
        titulo (string): Impresso no bloco da esquerda
        valor (Any): Impresso no bloco da direita
        nome_folha (openpyxl.worksheet.worksheet.Worksheet): deve estar no formato load_workbook(nome_arquivo)["nome_folha"]
    """    
    nome_folha.cell(row = linha_inicial,column = coluna_inicial, value = titulo)
    nome_folha.merge_cells(start_row = linha_inicial, start_column = coluna_inicial, end_row = linha_inicial+1, end_column = coluna_inicial +  1)
    nome_folha.cell(row = linha_inicial,column = coluna_inicial+2, value = valor)
    nome_folha.merge_cells(start_row = linha_inicial, start_column = coluna_inicial+2, end_row = linha_inicial+1, end_column = coluna_inicial +  3)

def criar_resumo(nome_folha,valor_total,valor_ultima_cotacao,quantidade,linha_inicial,coluna_inicial):
    """Cria os blocos de resumo do ativo

    Args:
        nome_folha (openpyxl.worksheet.worksheet.Worksheet): deve estar no formato load_workbook(nome_arquivo)["nome_folha"]
        valor_total (float): deve estar no formato 9999.99
        valor_ultima_cotacao (float): deve estar no formato 9999.99
        quantidade (float): deve estar no formato 9999.99
        linha_inicial (int): linha em que começa o bloco de resumos
        coluna_inicial (int): coluna em que começa o bloco de resumos
    """    
    #Cria os blocos do Total Anual
    linha_total = linha_inicial +1
    criar_blocos(linha_total,coluna_inicial,"Total Atual",valor_total,nome_folha)
    #Cria os blocos de Qtd. Ativos
    linha_quantidade = linha_inicial + 4
    criar_blocos(linha_quantidade,coluna_inicial,"Qtd. Ativos",quantidade,nome_folha)
    #Cria o bloco de Última Cotação
    linha_ultima_cotacao = linha_inicial + 7
    criar_blocos(linha_ultima_cotacao,coluna_inicial,"Última Cotação",valor_ultima_cotacao,nome_folha)

def criar_corpo_carteira(nome_folha,dicionario_ativos):
    """Cria todas as tabelas por cada ativo

    Args:
        nome_folha (openpyxl.worksheet.worksheet.Worksheet): deve estar no formato load_workbook(nome_arquivo)["nome_folha"]
        dicionario_ativos (dictionary): deve estar no formato {ativo1:qtd1,ativo2:qtd2,...}
    """    
    #Puxa o dicionario de ativos e seus dataframes
    cotacao = cotacao_semana(dicionario_ativos)
    #Define a linha que termina o cabeçalho
    linha_inicial = 9
    for ativo,data_frame in cotacao.items():
        criar_tabela_ativo(nome_folha,ativo,data_frame,linha_inicial,3)
        criar_resumo(nome_folha,1,1,1,linha_inicial,12)
        #Define o começo do próximo bloco
        linha_inicial += 12

def carteira(nome_arquivo,dicionario_ativos):
    """Realiza todas as modificações na folha carteira

    Args:
        nome_arquivo (string): deve estar no formato("nome_arquivo.xlsx")
        dicionario_ativos (dictionary): deve estar no formato {ativo1:qtd1,ativo2:qtd2,...}
    """    
    #Abre a planilha
    planilha = load_workbook(nome_arquivo)
    #Seleciona a folha da carteira
    folha_carteira = planilha["Carteira"]
    criar_cabecalho(folha_carteira)
    criar_corpo_carteira(folha_carteira,dicionario_ativos)
    #Salva o arquivo
    planilha.save(nome_arquivo)

nome_arquivo = "Teste.xlsx"
dicionario_teste = {"KO":"98987","MGLU3.SA":"1000000"}
criar_planilha(nome_arquivo)
carteira(nome_arquivo,dicionario_teste)
