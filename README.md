### RSLMaker
Programa para axuliar uma revisão sistemática. 

### Entrada: Um parameter .txt com diversos variáveis e, em especial, o diretorios de arquivos .bib
### Saída: Arquivo .xls

Esse programa ler arquivos .bib  e exporta para um arquivo de excel com uma pré-formatação e pré-filtragem (que será explicada nesse documento). Esses arquivos .bib são gerados através de buscas em bases de dados de pesquisa científica (como IEEE Xplorer, ACM Digital e Elsevier). Durante a leitura desses arquivos, o RSLMaker executa algumas filtragens conforme é solicitado. Por exemplo, é possível solicitar que artigos duplicados, ou seja, que apareçam em mais de um .bib, sejam vetados de aparecer mais de uma vez no excel.

O excel exportado pelo programa deve ser posteriormente formatado pelo usuário com o intutito de compartilha-lo entre os outros autores da revisão sistemática. Essa formatação fica a critério do usuário e a estrutura padrão dada como resposta do programa é uma estrutura básica e simples.

Por fim, detacamos que, nessa primeira versão do RSLMaker, os parâmetros estão dispostos como constantes no código fonte. Portanto, é preciso alterá-los antes de executar o mesmo. Essas constantes estão descritas no arquivo parameters.txt.


### Detalhes técnicos: 
O projeto foi codificado de forma simples, mediante sua simplicidade. A proposta é que possa ter contribuições de outros programadores. Para isso é necessário entrar em contato com o autor principal que de antemão garante dá os créditos devidos aos colaboradores. Esse projeto é gratuíto e deve-se sempre ser assim.


### Exemplo de execucao: 
python.exe main.py parameters.txt


### Exemplo de arquivos com patrametros:
diretorios = "../dadosIEEE/", "../dadosACM/"

saida = "../demo.xlsx"

titulos_proibidos = "survey", "review"

tipos_proibidos = "Book Chapter", "Conference Review", "Review", "inbook"

numero_minimo_paginas = -1

numero_maximo_paginas = 100

imprimir_artigos_reprovados = False


### Video exemplo de uso:
