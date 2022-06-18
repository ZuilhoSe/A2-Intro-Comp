# A2-Intro-Comp
Trabalho de A2 do curso de Introdução a Computação: scrapping de uma página web para avaliação de uma carteira financeira.

## Requisitos
É necessário ter as seguintes bibliotecas:
* tkinter
* openpyxl   
* pandas
* validators
* qrcode
* requests
* datetime 
* os

## Carteiras para utilização
* [https://zuilhose.github.io/A2-Intro-Comp/index.html](https://zuilhose.github.io/A2-Intro-Comp/index.html)
* [https://zuilhose.github.io/A2-Intro-Comp/carteira2.html](https://zuilhose.github.io/A2-Intro-Comp/carteira2.html)
* [https://zuilhose.github.io/A2-Intro-Comp/carteira3.html](https://zuilhose.github.io/A2-Intro-Comp/carteira3.html)

## Utilização
* Para realizar a consulta da carteira financeira basta executar o app.py, que executa um interface para inserção do link para a carteira e o nome desejado para o arquivo '.xlsx' de saída. 
* A execuçãdo do código é bem demorada, devido a forma como o yfinance trabalha, por isso, após fazer a requisição é necessário esperar algum tempo para ter o resultado.
* O arquivo de saída '.xlsx' possuí duas abas, uma com os valores da última semana da carteira e outra com os gráficos relativos aos ativos na carteira.

*obs_1: O arquivo de saída é gerado na pasta onde se encontra o arquivo 'app.py'*

*obs_2: Após iniciar o programa, a interface gráfica do tkinter informa que não está respondendo, porém, basta esperar o script acabar de rodar. Favor não fechar a janela*
