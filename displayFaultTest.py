__author__ = 'chenming'
import xlrd
import sys
from xml.parsers import expat


TranslationList = "E:\\python\\smallTest\\workstateTest.xlsx"
outputFile = 'e:\\python\\smallTest\\PYtest.exp'

workbook= xlrd.open_workbook(TranslationList)
mysheet=workbook.sheet_by_name("Translations")


SetIDRow = 1
IdCol = 0
StartRow = 2
StartCol = 1

CommentSymbol = '##'
languageSets =dict()

def cell2Str(cell):# cell convert to str type
    result='';
    if type(cell.value)==float:
        result = str(int(cell.value))
    elif cell.value >= u'\u4e00' and cell.value<=u'\u9fa5':# if it is chinese character
        result = str(cell.value.encode('GB2312'))
    else:
        result = str(cell.value)
    return result

for col in range(StartCol,mysheet.ncols):   #prepare the language sets
    languageSets[mysheet.cell_value(SetIDRow,col)] = "IF LanguageSet = " + str(int(mysheet.cell(SetIDRow,col).value))  +" THEN\n"
    languageSets[mysheet.cell_value(SetIDRow,col)] += "\tCASE " + str(mysheet.cell(0,IdCol).value)+ " OF\n"
    for line in range(StartRow,mysheet.nrows):
        languageSets[mysheet.cell_value(SetIDRow,col)] += "\t" + cell2Str(mysheet.cell(line,IdCol)) +" :\n"
        languageSets[mysheet.cell_value(SetIDRow,col)] += "\t\twork_display := '" + cell2Str(mysheet.cell(line,col))+ "';\n"

    languageSets[mysheet.cell_value(SetIDRow,col)] += "\tEND_CASE ;\n" + "END_IF\n"


#generate the declaration of the language wordids
declaration = "(*auto generated language definition file DO NOT alter manually *)\n"
declaration += "VAR\n"

outfile=open(outputFile,"w")

outfile.write( """(* @NESTEDCOMMENTS := 'Yes' *)
(* @PATH := '' *)
(* @OBJECTFLAGS := '0, 8' *)
(* @SYMFILEFLAGS := '2048' *)
PROGRAM Language
""")
outfile.write(declaration)
outfile.write( """END_VAR
(* @END_DECLARATION := '0' *)
""")
outfile.write("(*auto generated language definition file DO NOT alter manually *)\n")

for languageSet in languageSets:

    outfile.write( languageSets[languageSet])

outfile.close()



