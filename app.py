from excelToWordTool import read_excel_file, create_and_fill_word_file, create_pdf_files

excelFileName: str = 'example.xlsx'
excelSheetName: str = 'exampleSheet'
rowNumberForKeys: int = 9
excelColumnRangeWithData: str = 'E:G'
excelNRowsNumbersWithData: int = 4
docTemplateName: str = 'Hello.docx'
isPDFNeeded: bool = True
listOfDocTemplateVariables: list = ['name', 'id', 'password']  # must be in same order as keys in header of the Excel file
finalDocName: str = listOfDocTemplateVariables[0]  # prefer is to choose value of some template variables

excelDataToListOfDicts = read_excel_file(excelFileName,
                                         excelSheetName,
                                         rowNumberForKeys,
                                         excelColumnRangeWithData,
                                         excelNRowsNumbersWithData)

docsFiles = create_and_fill_word_file(excelDataToListOfDicts,
                                      docTemplateName,
                                      listOfDocTemplateVariables,
                                      finalDocName)

if isPDFNeeded:
    create_pdf_files(docsFiles)

