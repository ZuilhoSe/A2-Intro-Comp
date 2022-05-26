from openpyxl import Workbook, load_workbook

#Cria o arquivo excel
def criar_arquivo(nome_arquivo):
    planilha = Workbook()
    planilha.save(nome_arquivo)

#Coloca as folhas nomeadas
def criar_folhas(nome_arquivo):
    #Abre o arquivo
    planilha = load_workbook(nome_arquivo)
    folha1 = planilha.active
    #Criar Folhas nomeadas
    folha1.title = "Carteira"
    planilha.create_sheet("Estatísticas", 1)
    planilha.create_sheet("YFINANCE", 2)
    #Salva a planilha
    planilha.save(nome_arquivo)

#Cria o a planilha base
def criar_planilha(nome_carteira):
    nome_arquivo = nome_carteira +".xlsx"
    criar_arquivo(nome_arquivo)
    criar_folhas(nome_arquivo)

criar_planilha("Teste")
