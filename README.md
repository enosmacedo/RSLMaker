### RSLMaker
Programa para axuliar uma revisão sistemática. 

Esse programa ler arquivos .bib e exporta para um arquivo de excel com analises preliminares (abaixo). Esses arquivos .bib são gerados a partir de resultados de buscas de strings em bases de dados de pesquisa científica (como IEEE Explorer, ACM Digital e Elsevier). Durante a leitura desses arquivos, o RSLMaker executa algumas filtragens conforme é solicitado a ele para impedir que esses arquivos entrem no resultado final da planilha (arquivo excel). Por exemplo, é possível solicitar que artigos duplicados, ou seja, que apareçam em mais de um .bib, sejam vetados de aparecer mais de uma vez no excel.

O exel exportado pode ser formato pelo usuário com o intutito de compartilha-lo entre os outros autores da revisão sistemática. Essa formatação fica a critério do usuário e a estrutura padrão dada como resposta do programa é uma estrutura básica e simples.

Por fim, detacamos que: nessa primeira versão do RSLMaker, os parâmetros estão dispostos como constantes no código fonte. Portanto, é preciso alterá-los antes de executar o mesmo. Essas constantes estão descritas no arquivo parameters.txt.


### Exemplo de execucao: 
python.exe main.py parameters.txt

### Exemplo de arquivos com patrametros:
arquivos = "pubmed.bib", "outroarquivo.bib"

saida = "../demo.xlsx"

titulos_proibidos = "survey", "review"

tipos_proibidos = "Book Chapter", "Conference Review", "Review", "inbook"

numero_minimo_paginas = -1

numero_maximo_paginas = 100

imprimir_artigos_reprovados = False
