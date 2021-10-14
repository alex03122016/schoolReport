import docx

def analyzeDocX(inputDocx):

    #open .docx
    print(inputDocx)
    outputDocx = "saved-"+inputDocx
    savepath= outputDocx
    doc = docx.Document(inputDocx)

    tableDocx = doc.tables[0]
    for r in range(0,len(tableDocx.rows)):
        for c in range(0,len(tableDocx.columns)):
            sourceValue = "r"+str(r)+ "c"+str(c)
            targetCell = tableDocx.cell(r, c)
            targetCellValue = targetCell.text
            targetCell.text = str(sourceValue)

    tableDocx = doc.tables[1]
    for r in range(0,len(tableDocx.rows)):
        for c in range(0,len(tableDocx.columns)):
            sourceValue = "r"+str(r)+ "c"+str(c)
            targetCell = tableDocx.cell(r, c)
            print(len(tableDocx.rows))
            targetCellValue = targetCell.text
            #write data  to .docxf file
            targetCell.text = str(sourceValue)


    #save .docX
    doc.save(savepath)
    print("Was saved in", savepath)


if __name__ == "__main__":
    docxFile = "Z 420 - Notenzeugnis für Schülerinnen und Schüler mit dem Förderbedarf Lernen (01.21)-1.docx"
    analyzeDocX(docxFile)
