from openpyxl import Workbook, load_workbook

def criar_arquivo(nome_arquivo):
    """Cria o arquivo excel

    Args:
        nome_arquivo (string): deve estar no formato "nome_arquivo.xlsx"
    """    
    
    planilha = Workbook()
    planilha.save(nome_arquivo)


def criar_folhas(nome_arquivo):
    """Coloca as folhas nomeadas

    Args:
        nome_arquivo (string): deve estar no formato "nome_arquivo.xlsx"
    """
    
    #Abre o arquivo
    planilha = load_workbook(nome_arquivo)
    folha1 = planilha.active
    #Criar Folhas nomeadas
    folha1.title = "Carteira"
    planilha.create_sheet("Estatísticas", 1)
    #Salva a planilha
    planilha.save(nome_arquivo)

def link_folhas(nome_arquivo):
    """Cria o Link entre as duas folhas

    Args:
        nome_arquivo (string): deve estar no formato "nome_arquivo.xlsx"
    """
    
    #Abre o arquivo
    planilha = load_workbook(nome_arquivo)
    #Seleciona as folhas
    carteira = planilha["Carteira"]
    estatisticas = planilha["Estatísticas"]
    #Cria as strings dos links
    link_para_estatisticas = f"=HYPERLINK(\"[{nome_arquivo}]Estatísticas!A1\",\"Ir para Estatísticas ->\")"
    link_para_carteira = f"=HYPERLINK(\"[{nome_arquivo}]Carteira!A1\",\"Ir para Carteira ->\")"
    #Adiciona o link a celula A1 de cada folha
    carteira.merge_cells("A1:C1")
    carteira["A1"] = link_para_estatisticas
    estatisticas.merge_cells("A1:C1")
    estatisticas["A1"] = link_para_carteira
    #Salva a planilha
    planilha.save(nome_arquivo)

def criar_planilha(nome_arquivo):
    """Cria o a planilha base

    Args:
        nome_arquivo (string): deve estar no formato "nome_arquivo.xlsx"
    """
    criar_arquivo(nome_arquivo)
    criar_folhas(nome_arquivo)
    link_folhas(nome_arquivo)

criar_planilha("teste.xlsx")