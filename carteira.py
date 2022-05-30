from openpyxl import Workbook, load_workbook
from cotacao import cotacao_semana
from openpyxl.utils.dataframe import dataframe_to_rows
from criar_excel import criar_planilha

def criar_cabecalho(nome_arquivo):
    """Cria o cabeçalho da planilha

    Args:
        nome_arquivo (string): deve estar no formato "nome_arquivo.xlsx"
    """    
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

def criar_tabela_ativo(nome_arquivo,ativo,data_frame,linha_inicial,coluna_inicial):
    """Cria na planilha a tabela por ativo

    Args:
        nome_arquivo (string): deve estar no formato "nome_arquivo.xlsx"
        ativo (string): nome do ativo que será o título
        data_frame (dataframe): dataframe que se tornará a tabela
        linha_inicial (int): linha em que começa o bloco da tabela
        coluna_inicial (int): coluna em que começa o bloco
    """    
    #Abre o arquivo
    planilha = load_workbook(nome_arquivo)
    #Seleciona a folha
    carteira = planilha["Carteira"]
    #Cria o título do ativo
    linha_titulo = linha_inicial + 1
    carteira.cell(row = linha_titulo,column = coluna_inicial, value = ativo)
    carteira.merge_cells(start_row = linha_titulo, start_column = coluna_inicial, end_row = linha_titulo+1, end_column = coluna_inicial +  7)
    #Define a linha inicial da tabela
    linha_tabela = linha_inicial + 4 
    carteira.cell(row = linha_tabela,column = coluna_inicial, value = None)
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
            carteira.cell(row = linha_tabela,column = coluna_item, value = item_coluna)
            coluna_item += 1
        linha_tabela += 1    
    #Salva o arquivo
    planilha.save(nome_arquivo)

def criar_corpo_carteira(nome_arquivo,dicionario_ativos):
    """Cria todas as tabelas por cada ativo

    Args:
        nome_arquivo (string): deve estar no formato "nome_arquivo.xlsx"
        dicionario_ativos (dictionary): deve estar no formato {ativo1:qtd1,ativo2:qtd2,...}
    """    
    #Puxa o dicionario de ativos e seus dataframes
    cotacao = cotacao_semana(dicionario_ativos)
    #Define a linha que termina o cabeçalho
    linha_inicial = 9
    for ativo,data_frame in cotacao.items():
        criar_tabela_ativo(nome_arquivo,ativo,data_frame,linha_inicial,3)
        #Define o começo do próximo bloco
        linha_inicial += 12


nome_arquivo = "Teste.xlsx"
criar_planilha(nome_arquivo)
criar_cabecalho(nome_arquivo)
dicionario_teste = {"MGLU3.SA":"1000000", "KO":"98987"}
criar_corpo_carteira(nome_arquivo,dicionario_teste)
