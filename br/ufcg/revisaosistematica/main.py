import requests
import json
import bibtexparser
from bibtexparser.bparser import BibTexParser
import xlsxwriter


FORBIDDEN_WORDS_TITLE = ["survey", "review"]
DISNABLE_TYPES = ["Book Chapter", "Conference Review", "Review", "inbook"]
# document_type Book Chapter - Conference Review  - Review - inproceedings
NUM_MIN_PAGES = 3

data_final = [];
data_reject = [];


def search(a1, msg, qnt_arquivos):
    global data_final
    global data_reject
    qnt_reject_title = 0;
    qnt_reject_types = 0;
    qnt_reject_duplicate = 0;
    qnt_reject_num_pages = 0;
    qnt_ok = 0;
    qnt_reject = 0;
    qnt_total = 0;

    data = []
    for i in range(0, qnt_arquivos):
        bib_database = None
        parser = BibTexParser(common_strings=True)
        if i == 0:
            with open(a1+ ".bib", encoding="utf8") as bibtex_file:
                bib_database = bibtexparser.load(bibtex_file, parser=parser)
        else:
            with open(a1+"(" + str(i) + ").bib", encoding="utf8") as bibtex_file:
                bib_database = bibtexparser.load(bibtex_file, parser=parser)
        data = data + bib_database.entries

    for artigo in data:
        qnt_total = qnt_total + 1
        can_add = True;
        for disable_word in FORBIDDEN_WORDS_TITLE:
            if disable_word.lower() in artigo["title"].lower():
                qnt_reject_title = qnt_reject_title + 1
                can_add = False;
                break;

        if can_add:
            for disable_type in DISNABLE_TYPES:
                if (disable_type.lower() in artigo["ENTRYTYPE"].lower() ):
                    can_add = False;
                    qnt_reject_types = qnt_reject_types + 1
                    break

        if can_add:
            for artigo_ja_adicionado in data_final:
                if artigo["title"].casefold().replace(" ", "").lower() == artigo_ja_adicionado["title"].casefold().replace(" ", "").lower():
                    can_add = False;
                    qnt_reject_duplicate = qnt_reject_duplicate + 1
                    break;

        if can_add:
            for artigo_ja_adicionado in data_final:
                try:
                    if (artigo["document_type"].casefold().replace(" ", "").lower() == artigo_ja_adicionado["document_type"].casefold().replace(" ", "").lower()):
                        can_add = False;
                        qnt_reject_types = qnt_reject_types + 1
                        break
                except:
                    continue

        if can_add:
            try:
                if (int(artigo["numpages"])  < NUM_MIN_PAGES) :
                    can_add = False;
                    # print("Excluido por pagina - " +  msg + ": " + artigo["title"])
                    qnt_reject_num_pages = qnt_reject_num_pages + 1
            except:
                try:
                    inicial = int(artigo["pages"].split("-")[0])
                    final = int(artigo["pages"].split("-")[1])

                    if ((final - inicial + 1) < NUM_MIN_PAGES):
                        can_add = False;
                        # print("Excluido por pagina - " + msg + ": " + artigo["title"])
                        qnt_reject_num_pages = qnt_reject_num_pages + 1
                except:
                    print("Conferir pagina - " +  msg + ": " + artigo["title"])


        try:
            aux = {"title": artigo["title"], "abstract": artigo["abstract"],  "keywords": artigo["keywords"], "content_type": artigo["ENTRYTYPE"],"publication_year": artigo["year"], "plataforma": msg, "bibtex": str(artigo)}
        except:
            aux = {"title": artigo["title"], "abstract": ""                ,  "keywords": ""                , "content_type": artigo["ENTRYTYPE"], "publication_year": artigo["year"], "plataforma": msg, "bibtex": str(artigo)}
        if can_add:
            data_final = data_final + [aux]
            qnt_ok = qnt_ok + 1
        else:
            data_reject = data_reject + [aux]
            qnt_reject = qnt_reject + 1

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
    # Create an new Excel file and add a worksheet.
    workbook = xlsxwriter.Workbook('demo.xlsx')
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

    row = 2;
    for a in data:
        worksheet.write('A' + str(row), a['title'])
        worksheet.write('B' + str(row), a['content_type'])
        worksheet.write('C' + str(row), a['publication_year'])
        worksheet.write('D' + str(row), a['abstract'])
        worksheet.write('E' + str(row), a['keywords'])
        worksheet.write('F' + str(row), a['plataforma'])
        worksheet.write('G' + str(row), a['bibtex'])
        row = row + 1;

    workbook.close()
    return


def main():
    search("../../../files/acm", "acm", 1)
    search("../../../files/scopus", "scopus", 1)
    search("../../../files/ieee", "ieee", 1)

    # search("../../../filesmarcus/acm", "acm", 1)
    # search("../../../filesmarcus/scopus", "scopus", 1)
    # search("../../../filesmarcus/ieee", "ieee", 2)


    print ("aprovados final: " + str(len(data_final)))
    create_xlms(data_final)

    print("reprovados final: " + str(len(data_reject)))
    index = 0;
    for a in data_reject:
        index = index + 1;
        print(index, " -->", a["title"])
    return


if __name__ == "__main__":
    main()

