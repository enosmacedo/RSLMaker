# RSLMaker
Programa para axuliar a revisão sistemática. 

O programa ler arquivos .bib e exporta para um arquivo de excel com analises preliminares (abaixo). Depois da exportação, o autor pode pegar o excel e formatá-lo condicionalmente para compartilhar entre os outros autores da revisão sistemática. Isso é critério de cada autor, mas, se precisar, é só me comunicar que eu tenho planilhas já formatadas.


Exemplo de execucao: 
python.exe main.py parameters.txt


Exemplo de arquivos com patrametros:


arquivos = "pubmed.bib", " "outroarquivo.bib""

saida = "../demo.xlsx"

titulos_proibidos = "survey", "review"

tipos_proibidos = "Book Chapter", "Conference Review", "Review", "inbook"

numero_minimo_paginas = -1

numero_maximo_paginas = 100

imprimir_artigos_reprovados = False
