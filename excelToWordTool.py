from pandas import read_excel
from docxtpl import DocxTemplate
import comtypes.client
import os


def read_excel_file(excelName, sheetName, header, columns, rows):
    data = read_excel(excelName, sheet_name=sheetName, header=header,
                      usecols=columns, nrows=rows,
                      engine="openpyxl")

    return data.to_dict('records')


def create_and_fill_word_file(excelData, docTemplateName, contextKeys, docName):
    docNames = []
    for item in excelData:
        valuesToFill = list(item.values())

        context = {key: value
                   for key, value in zip(contextKeys, valuesToFill)
                   }
        try:
            fileNameDoc = f"{context[docName]}.docx"
        except KeyError:
            print("Given context key to name doc file doesn't exists. File was named using key itself.")
            fileNameDoc = f"{docName}.docx"
        doc = DocxTemplate(docTemplateName)
        doc.render(context)
        doc.save(fileNameDoc)
        docNames.append(fileNameDoc)
    return docNames


def create_pdf_files(docFileNames):
    for doc in docFileNames:
        fileNamePdf = doc.replace(".docx", ".pdf")
        wdFormatPDF = 17
        word = comtypes.client.CreateObject('Word.Application')
        doc = word.Documents.Open(os.getcwd() + '\\' + doc)
        doc.SaveAs(os.getcwd() + '\\' + fileNamePdf, FileFormat=wdFormatPDF)
        doc.Close()
        word.Quit()
