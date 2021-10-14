import docx
from schoolReport import getNamesGrades, coordinates, testDictionary

def writeSchoolReport(inputDocx):
    xlsxFile = "/home/alex/schoolReport/Notenliste-11-10-2021_04-02-21.xlsx"
    #names, allGrades = getNamesGrades.getNamesGrades(xlsxFile)
    names, allGrades = testDictionary.testDictionary(xlsxFile)

    coordinatesDocX = coordinates.coordinates()

    #get name
    pupil = 0
    for name in names:
        #open .docx
        print(inputDocx)
        outputDocx = "output/"+name+"-saved-"+inputDocx
        savepath= outputDocx
        doc = docx.Document(inputDocx)

        #get Value from Source
        for schoolSubject in allGrades:
            note = allGrades[schoolSubject]["Gesamtnote"][pupil]
            print(schoolSubject, note )

            #get Coordinates
            for i in range(len(coordinatesDocX)):

                if schoolSubject == coordinatesDocX[i][0]:
                    c = coordinatesDocX[i][3]
                    r = coordinatesDocX[i][2]

                    print("coordinates of:  "+schoolSubject+" column: "+str(c)+", row: " +str(r))
                    tableDocx = doc.tables[coordinatesDocX[i][4]]
                    targetCell = tableDocx.cell(r, c)

                    #write to target
                    targetCell.text = str(note)
        pupil += 1



        #save to target .docX
        doc.save(savepath)
        print("Was saved in", savepath)


if __name__ == "__main__":

    docxFile = "Z 420 - Notenzeugnis für Schülerinnen und Schüler mit dem Förderbedarf Lernen (01.21)-1.docx"
    writeSchoolReport(docxFile)
