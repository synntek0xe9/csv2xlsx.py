import xlwings
import sys
import requests



def copyTemplate(template="template.xlsx", newFile="newfile.xlsx"):


    fin = open(template,"rb")
    fout = open(newFile,"wb")
    fout.write(fin.read()); fin.close(); fout.close()


def readCsv(text,separator=","):

    lines = text.split('\n')

    data = []
    for line in lines:
        row = line.split(separator)
        data.append(row)

    if len(data[-1]) < len(data[0]):
        data = data[:-1]

    return data




if __name__ == "__main__":

    
    # default value that is overriden later
    url = "https://www.stats.govt.nz/assets/Uploads/Annual-enterprise-survey/Annual-enterprise-survey-2023-financial-year-provisional/Download-data/annual-enterprise-survey-2023-financial-year-provisional-size-bands.csv"
    keepOpen = False
    separator = ";"

    if "-u" in sys.argv:
        url = sys.argv[sys.argv.index("-u")+1]
        separator = ","
    if "--keep-open" in sys.argv:
        keepOpen = True


    text = requests.get(url).text

    copyTemplate(template="csvTemplate.xlsx", newFile="fromCsv.xlsx")

    data = readCsv(text, separator=separator)

    book = xlwings.Book("fromCsv.xlsx")
    dataSheet = book.sheets['dane']
    dataSheet.range("A1").value = data 
    # when changing range value - range automatically expands in right/down direction
    # when changing range value - new value has to be array (or 2d array) 

    book.save()
    if not keepOpen:
        book.close()