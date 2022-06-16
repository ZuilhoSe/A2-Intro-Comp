from tkinter import *
import validators
import scrapper


class Application:
    def __init__(self, master=None):
        self.fontePadrao = ("Arial", "10")
        self.primeiroContainer = Frame(master)
        self.primeiroContainer["pady"] = 10
        self.primeiroContainer.pack()

        self.segundoContainer = Frame(master)
        self.segundoContainer["padx"] = 20
        self.segundoContainer.pack()

        self.terceiroContainer = Frame(master)
        self.terceiroContainer["padx"] = 20
        self.terceiroContainer.pack()

        self.quartoContainer = Frame(master)
        self.quartoContainer["pady"] = 20
        self.quartoContainer.pack()

        self.titulo = Label(self.primeiroContainer, text="Robô de Avaliação de Portfólio de Investimentos")
        self.titulo["font"] = ("Arial", "10", "bold")
        self.titulo.pack()

        self.urlLabel = Label(self.segundoContainer,text="URL", font=self.fontePadrao)
        self.urlLabel.pack(side=LEFT)

        self.url = Entry(self.segundoContainer)
        self.url["width"] = 30
        self.url["font"] = self.fontePadrao
        self.url.pack(side=LEFT)

        self.iniciar = Button(self.quartoContainer)
        self.iniciar["text"] = "Iniciar"
        self.iniciar["font"] = ("Calibri", "8")
        self.iniciar["width"] = 12
        self.iniciar["command"] = self.verificaURL
        self.iniciar.pack()

        self.mensagem = Label(self.quartoContainer, text="", font=self.fontePadrao)
        self.mensagem.pack()

    #Método verificar url
    def verificaURL(self):
        self.url_coletada = self.url.get()
        if validators.url(self.url_coletada) == True:
            self.mensagem["text"] = "URL válida, dando prosseguimento..."
            self.realizar_scrapper()
            self.mensagem["text"] = str(self.carteira)
        else:
            self.mensagem["text"] = "URL INVÁLIDA! Certifique-se de que a URL está correta!"

    #Método para realização do scrapper;
    def realizar_scrapper(self):
        soup = scrapper.ler_carteira(self.url_coletada)
        acoes = scrapper.buscar_acoes(soup)
        moedas = scrapper.buscar_moedas(soup)
        lista = scrapper.juntar_moedas_acoes(acoes, moedas)
        self.carteira = scrapper.saida(lista)

    #Método que executa o módulo de buasca de cotações;
    def buscar_cotacoes(self):
        pass

#Execução do Aplciativo
root = Tk()
Application(root)
root.mainloop()

