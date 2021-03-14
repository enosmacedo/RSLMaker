import sys
import json
import bibtexparser
from bibtexparser.bparser import BibTexParser
import xlsxwriter

FILE_NAME_PARAMETER = "arquivos"
OUTPUT_NAME_PARAMETER = "saida"
FORBIDDEN_WORDS_TITLE_PARAMETER = "titulos_proibidos"
DISABLE_TYPES_PARAMETER = "tipos_proibidos"
NUM_MIN_PAGES_PARAMETER = "numero_minimo_paginas"
NUM_MAX_PAGES_PARAMETER= "numero_maximo_paginas"
PRINT_REPROVED_ARTICLES_PARAMETER = "imprimir_artigos_reprovados"

forbidden_words_title = []
disable_types = []
num_min_pages = -1
num_max_pages = 10000
print_reproved_articles = False
output_file = ""

data_final = [];
data_reject = [];


def search(a1, msg, qnt_arquivos):
    global data_final
    global data_reject
    global forbidden_words_title
    global disable_types
    global num_min_pages
    global num_max_pages
    global print_reproved_articles
    global output_file

    qnt_reject_title = 0;
    qnt_reject_types = 0;
    qnt_reject_duplicate = 0;
    qnt_reject_num_pages = 0;
    qnt_reject_library = 0;
    qnt_ok = 0;
    qnt_reject = 0;
    qnt_total = 0;
    cause_exclusion = ""

    data = []
    for i in range(0, qnt_arquivos):
        bib_database = None
        parser = BibTexParser(common_strings=True)
        if i == 0:
            with open(a1, encoding="utf8") as bibtex_file:
                bib_database = bibtexparser.load(bibtex_file, parser=parser)
        else:
            with open(a1+"(" + str(i) + ")", encoding="utf8") as bibtex_file:
                bib_database = bibtexparser.load(bibtex_file, parser=parser)
        data = data + bib_database.entries

    for idx, artigo in enumerate(data):
        qnt_total = qnt_total + 1
        can_add = True;
        for disable_word in forbidden_words_title:
            try:
                if disable_word.lower() in artigo["title"].lower():
                    qnt_reject_title = qnt_reject_title + 1
                    can_add = False;
                    cause_exclusion = "Por titulo"
                    break
            except:
                pass
        if can_add:
            for disable_type  in disable_types:
                try:
                    if (disable_type.lower() in artigo["ENTRYTYPE"].lower() ):
                        can_add = False;
                        qnt_reject_types = qnt_reject_types + 1
                        cause_exclusion = "Por entrytype"
                        break
                except:
                    pass
        if can_add:
            for artigo_ja_adicionado in data_final:
                try:
                    if artigo["title"].casefold().replace(" ", "").lower() == artigo_ja_adicionado["title"].casefold().replace(" ", "").lower():
                        can_add = False;
                        qnt_reject_duplicate = qnt_reject_duplicate + 1
                        cause_exclusion = "Por titulo - reptido"
                        break;
                except:
                    pass

        if can_add:
            for artigo_ja_adicionado in data_final:
                try:
                    if (artigo["document_type"].casefold().replace(" ", "").lower() == artigo_ja_adicionado["document_type"].casefold().replace(" ", "").lower()):
                        can_add = False;
                        qnt_reject_types = qnt_reject_types + 1
                        cause_exclusion = "Por tipo documento"
                        break
                except:
                    pass

        if can_add:
            try:
                if (int(artigo["numpages"])  < num_min_pages) :
                    can_add = False;
                    qnt_reject_num_pages = qnt_reject_num_pages + 1
                    cause_exclusion = "Por numero paginas minimo"
            except:
                try:
                    inicial = int(artigo["pages"].split("-")[0])
                    final = int(artigo["pages"].split("-")[-1])
                    if ((final - inicial + 1) < num_min_pages):
                        can_add = False;
                        qnt_reject_num_pages = qnt_reject_num_pages + 1
                        cause_exclusion = "Por numero paginas minimo"
                except:
                    pass

        if can_add:
            try:
                if (int(artigo["numpages"])  > num_max_pages) :
                    can_add = False;
                    qnt_reject_num_pages = qnt_reject_num_pages + 1
                    cause_exclusion = "Por numero paginas maximo"
            except:
                try:
                    inicial = int(artigo["pages"].split("-")[0])
                    final = int(artigo["pages"].split("-")[-1])
                    if ((final - inicial) > num_max_pages):
                        can_add = False;
                        qnt_reject_num_pages = qnt_reject_num_pages + 1
                        cause_exclusion = "Por numero paginas maximo"
                except:
                    pass

        aux = {"title": artigo["title"]}
        try:
           aux["abstract"] = artigo["abstract"]
        except:
           aux["abstract"] = " "
        try:
           aux["keywords"] = artigo["keywords"]
        except:
           aux["keywords"] = " "
        try:
           aux["content_type"] = artigo["ENTRYTYPE"]
        except:
           aux["content_type"] = " "
        try:
           aux["publication_year"] = artigo["year"]
        except:
           aux["publication_year"] = " "
        aux["plataforma"] = msg
        aux["bibtex"] = str(artigo)

        if can_add:
            data_final = data_final + [aux]
            qnt_ok = qnt_ok + 1
        else:
            data_reject = data_reject + [aux]
            qnt_reject = qnt_reject + 1
            if print_reproved_articles:
                print(cause_exclusion + " - " + str(idx) + " - " +  msg + ": " + artigo["title"])

    print("total "     + msg + " no total inicio (1): " + str(len(data)))
    print("aprovados " + msg + " no total inicio (2): " + str(qnt_total))
    print("aprovados " + msg + " no final: " + str(qnt_ok))
    print("reprovados "+ msg + " total: " + str(qnt_reject))
    print("reprovados "+ msg + " por titulo: " + str(qnt_reject_title))
    print("reprovados "+ msg + " por tipo: " + str(qnt_reject_types))
    print("reprovados "+ msg + " por duplicacao: " + str(qnt_reject_duplicate))
    print("reprovados "+ msg + " por num pages: " + str(qnt_reject_num_pages))
    print("")
    return data


def create_xlms(data):
    global output_file
    # Create an new Excel file and add a worksheet.
    workbook = xlsxwriter.Workbook(output_file)
    worksheet = workbook.add_worksheet()

    # Add a bold format to use to highlight cells.
    bold = workbook.add_format({'bold': True})

    # Text with formatting.
    worksheet.write('A1', 'title', bold)
    worksheet.write('B1', 'content_type', bold)
    worksheet.write('C1', 'publication_year', bold)
    worksheet.write('D1', 'abstract', bold)
    worksheet.write('E1', 'keywords', bold)
    worksheet.write('F1', 'plataforma', bold)
    worksheet.write('G1', 'bibtex', bold)

    row_initial_aux = 2;
    for a in data:
        worksheet.write('A' + str(row_initial_aux), a['title'])
        worksheet.write('B' + str(row_initial_aux), a['content_type'])
        worksheet.write('C' + str(row_initial_aux), a['publication_year'])
        worksheet.write('D' + str(row_initial_aux), a['abstract'])
        worksheet.write('E' + str(row_initial_aux), a['keywords'])
        worksheet.write('F' + str(row_initial_aux), a['plataforma'])
        worksheet.write('G' + str(row_initial_aux), a['bibtex'])
        row_initial_aux = row_initial_aux + 1;

    workbook.close()
    return


def main():
    global forbidden_words_title
    global disable_types
    global num_min_pages
    global num_max_pages
    global print_reproved_articles
    global output_file

    try: 
        fileParameters = sys.argv[1]
        fileObj = open(fileParameters)
    except: 
        mensagem_aviso = "Passe como parametro de execucao um arquivo (preferencialmente .txt) "
        mensagem_aviso = mensagem_aviso + "com todos os parametros de execucao. Nao eh possivel omitir parametros, ou seja, no arquivo tem " 
        mensagem_aviso = mensagem_aviso + "que ter todos os parametros, mesmo que sejam vazios - veja o arquivo parameters.txt como exemplo"
        print(mensagem_aviso)

    params = {}
    for line in fileObj:
        line = line.strip()
        key_value = line.split("=")
        if len(key_value) == 2:
            value_vector = key_value[1].split(",")
            aux = []
            for j in value_vector:
                j = j.strip(" ").strip("\"")
                aux.append(j)

            params[key_value[0].strip()] = aux

    try: 
        for  p in params[FORBIDDEN_WORDS_TITLE_PARAMETER]: 
            forbidden_words_title.append(p)
    except: 
        pass
    try:
        for  p in params[DISABLE_TYPES_PARAMETER]: 
            disable_types.append(p)
    except: 
        pass
    try:
        num_min_pages = int(params[NUM_MIN_PAGES_PARAMETER][0])
    except: 
        pass
    try:
        num_max_pages = int(params[NUM_MAX_PAGES_PARAMETER][0])
    except: 
        pass
    try:
        print_reproved_articles = list(map(lambda ele: ele == "True", params[PRINT_REPROVED_ARTICLES_PARAMETER]))[0]
    except: 
        pass
    try:
        output_file = params[OUTPUT_NAME_PARAMETER][0]
    except: 
        pass
    
    try:
        for  f in params[FILE_NAME_PARAMETER]: 
            aux = f.split("/")[-1].split(".")[0];
            search(f, aux, 1)
    except:
        print("Eh preciso adicionar os arquivos .bib")
   
    print ("aprovados final (após todas as análises): " + str(len(data_final)))
    create_xlms(data_final)
    
    print("reprovados final (após todas as análises): " + str(len(data_reject)))
    if print_reproved_articles:
        index = 0;
        for a in data_reject:
            index = index + 1;
            print(index, " -->", a["title"])
    return


if __name__ == "__main__":
    main()

