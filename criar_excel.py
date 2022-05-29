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
    #Salva a planilha
    planilha.save(nome_arquivo)

#Cria o Link entre as duas folhas
def link_folhas(nome_arquivo):
    #Abre o arquivo
    planilha = load_workbook(nome_arquivo)
    #Seleciona as folhas
    carteira = planilha["Carteira"]
    estatisticas = planilha["Estatísticas"]
    #Cria as strings dos links
    link_para_estatisticas = f"=HYPERLINK(\"[{nome_arquivo}]Estatísticas!A1\",\"Estatísticas\")"
    link_para_carteira = f"=HYPERLINK(\"[{nome_arquivo}]Carteira!A1\",\"Carteira\")"
    #Adiciona o link a celula A1 de cada folha
    carteira["A1"] = link_para_estatisticas
    estatisticas["A1"] = link_para_carteira
    #Salva a planilha
    planilha.save(nome_arquivo)

#Cria o a planilha base
def criar_planilha(nome_arquivo):
    criar_arquivo(nome_arquivo)
    criar_folhas(nome_arquivo)
    link_folhas(nome_arquivo)
