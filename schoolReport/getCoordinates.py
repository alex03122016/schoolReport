
def getCoordinates(schoolSubj, coordinatesDocX, doc):
    """ get Coordinates of "note" for print to docX """
    #get Coordinates
    for i in range(len(coordinatesDocX)):

        if schoolSubj == coordinatesDocX[i][0]:
            c = coordinatesDocX[i][3]
            r = coordinatesDocX[i][2]

            print("coordinates of:  "+schoolSubj+" column: "+str(c)+", row: " +str(r))
            tableDocx = doc.tables[coordinatesDocX[i][4]]
            targetCell = tableDocx.cell(r, c)

            return targetCell
