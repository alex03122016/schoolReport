import docx
import os
#from schoolReport import getNamesGrades, coordinates, testDictionary, writePupilDetailsToSchoolReport, testwritePupilDetailsToSchoolReport
from schoolReport import getCoordinates as getCo
from schoolReport import coordinates, getParticipation


def writeParticipationToSchoolReport(inputDocx,
                                    docxFile,
                                    nameOfPupil,
                                    Geburtsdatum="",
                                    save= "test",
                                    xlsxFile="",
                                    pupil=0):
    """ will overwrite the inputDocx File
    looks for #Vorname and #Geburtsdatum in docx file and replaces it
    with data given as variable nameOfPupil and Geburtsdatum"""


    #open .docx
    print(inputDocx)
    #outputDocx = "saved-"+inputDocx
    outputDocx = os.path.join(  os.path.expanduser('~'),
                                "schoolReport",
                                "output",
                                 save + docxFile)
    savepath= outputDocx
    doc = docx.Document(inputDocx)
    coordinatesDocX = coordinates.coordinates()
    names, allGrades = getParticipation.getParticipation(xlsxFile)

    Fehlzeiten = ["Stu","Stu_ue","v",]

    for zeit in Fehlzeiten:
        getCoordinates = getCo.getCoordinates(  schoolSubj=zeit,
                                            coordinatesDocX=coordinatesDocX,
                                            doc=doc)
        #get note from allGrades
        #note = allGrades["7a - Deu"]["Ø Klausur"][pupil]
        note = zeit
        participation = allGrades[zeit][pupil]

        print(zeit)
        targetCell = getCoordinates
        if targetCell != None:
            #write to target
            targetCell.text = str(participation)
        else:
            print("target cell is None")

    #save .docX
    #uncomment for integration in writeSchoolReport.py
    doc.save(inputDocx)
    print("Was saved in", inputDocx)

    #uncomment for testing
    #doc.save(outputDocx)
    #print("Was saved in", outputDocx)
    return

if __name__ == "__main__":
    xlsxFile = "/home/alex/schoolReport/Unterrichtsstunden-18-10-2021_15-15-57.xlsx"
    docxFileTest = "Z 420 - Notenzeugnis für Schülerinnen und Schüler mit dem Förderbedarf Lernen (01.21)-1.docx"
    inputDocxTest = os.path.join(os.path.expanduser('~'),"schoolReport", docxFileTest)
    nameOfPupilTest = "Alexander Tanck"
    GeburtsdatumTest = "26. September 1986"
    saveTest = "TestParticipation"

    writeParticipationToSchoolReport(inputDocx=inputDocxTest,
                                    docxFile=docxFileTest,
                                    nameOfPupil=nameOfPupilTest,
                                    Geburtsdatum=GeburtsdatumTest,
                                    save=saveTest )
