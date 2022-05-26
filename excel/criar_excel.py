from openpyxl import Workbook

#Cria a planilha
def criar_planilha(nome_da_planilha):
    planilha = Workbook()
    nome_arquivo = nome_da_planilha + ".xlsx"
    planilha.save(nome_arquivo)

criar_planilha("teste")
