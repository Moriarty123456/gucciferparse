import sys

sys.path.append("/home/david/Documents/ClintonJail/Guccifer2.0/Docs&Tools/Python/ALL_GUC_FILE_PROJECT/python/modules/DOC")

from binarytoolkit import *
#infile = "/home/david/Documents/ClintonJail/Guccifer2.0/Docs&Tools/Python/ALL_GUC_FILE_PROJECT/downloaded_docs/doc/old/briefing-norman-hsu.doc"

def ExtraPositions(infile):
    herx = OutputHex(infile)
    # get official location for root entry
    header = OutputHeader(herx)
    officialrootentryPosition = (int(ReverseHex(Header2FourBytes(header)[12]),16)+1)*512
    #Now search for all "Root Entries", and add to a list those that are not the official position.
    # 52006f006f007400200045006e00740072007900
    AllRootDirectories = PositionChunkList(herx, "52006f006f007400200045006e00740072007900",2048)
    AllRootDirectoriesPositions = []
    x = 0
    while x < len(AllRootDirectories):
        AllRootDirectoriesPositions.append(AllRootDirectories[x])
        x = x+2
    # If there are extra Roots, loop through them and pick out the extra
    extrarootpositions = []
    for y in AllRootDirectoriesPositions:
        if y != officialrootentryPosition:
            extrarootpositions.append(y)
        else:
            pass
    # return the extra positions list
    return extrarootpositions




def AlternatePositions(infile):
    exRootList = ExtraPositions(infile)
    PositionsList = []
    filename = infile.split("/")[-1]
    if len(exRootList) > 0:
        for x in exRootList:
            PositionsList.append("ALTERNATE_"+filename)
            PositionsList.append(x)
            herx = OutputHex(infile)
            alternaterootentryPosition = x

            alternateRootSlice = herx[2*alternaterootentryPosition:(2*alternaterootentryPosition)+2048]
            rtlist=RootTimes(Header2FourBytes(alternateRootSlice))
            PositionsList.append("ALTERNATERootTimes")
            PositionsList.append(rtlist)
            PositionsList.append("ALTERNATERootSlice")
            PositionsList.append(alternateRootSlice)
            # 1) RootSliceDocumentSummaryInfo
            if "05004400" in alternateRootSlice:
                try:
                    RootSliceDocumentSummaryInfoChunk = PositionChunkList(alternateRootSlice, "05004400", 256)
                except ValueError:
                    pass
            else:
                RootSliceDocumentSummaryInfoChunk = [0,0]
            if RootSliceDocumentSummaryInfoChunk[1] != 0:
                try:
                    thePosition = (int(ReverseHex(Header2FourBytes(RootSliceDocumentSummaryInfoChunk[1])[29]),16)+1)*512
                    if thePosition == 512:
                        RootDocSumInfoPosition = 0
                    else:
                        RootDocSumInfoPosition = thePosition
                except ValueError:
                    RootDocSumInfoPosition = 0
                try:
                    RootDocSumInfoPositionLength = int(ReverseHex(Header2FourBytes(RootSliceDocumentSummaryInfoChunk[1])[30]),16)
                except ValueError:
                    RootDocSumInfoPositionLength =0
            else:
                RootDocSumInfoPosition = 0
                RootDocSumInfoPositionLength = 0
        
            # 2) RootSliceData
            if "44006100" in alternateRootSlice:
                try:
                    RootSliceDataChunk = PositionChunkList(alternateRootSlice, "44006100", 256)
                except ValueError:
                    pass
            else:
                RootSliceDataChunk = [0,0]
            if RootSliceDataChunk[1] != 0:
                try:
                    thePosition = (int(ReverseHex(Header2FourBytes(RootSliceDataChunk[1])[29]),16)+1)*512
                    if thePosition == 512:
                        RootSliceDataPosition = 0
                    else:
                        RootSliceDataPosition = thePosition
                except ValueError:
                    RootSliceDataPosition = 0
                try:
                    RootSliceDataLength = int(ReverseHex(Header2FourBytes(RootSliceDataChunk[1])[30]),16)
                except ValueError:
                    RootSliceDataLength =0
            else:
                RootSliceDataPosition = 0
                RootSliceDataLength = 0
            # 3) 1Table

            if "31005400" in alternateRootSlice:
                try:
                    RootSlice1TableChunk = PositionChunkList(alternateRootSlice, "31005400", 256)
                except ValueError:
                    pass
            else:
                RootSlice1TableChunk = [0,0]
            if RootSlice1TableChunk[1] != 0:
                try:
                    thePosition = (int(ReverseHex(Header2FourBytes(RootSlice1TableChunk[1])[29]),16)+1)*512
                    if thePosition == 512:
                        RootSlice1TablePosition = 0
                    else:
                        RootSlice1TablePosition = thePosition
                except ValueError:
                    RootSlice1TablePosition = 0
                try:
                    RootSlice1TableLength = int(ReverseHex(Header2FourBytes(RootSlice1TableChunk[1])[30]),16)
                except ValueError:
                    RootSlice1TableLength =0
            else:
                RootSlice1TablePosition = 0
                RootSlice1TableLength = 0
        
            #4) Word Document
    
            if "57006f0072006400" in alternateRootSlice:
                try:
                    RootSliceWordDocumentChunk = PositionChunkList(alternateRootSlice, "57006f0072006400", 256)
                except ValueError:
                    pass
            else:
                RootSliceWordDocumentChunk = [0,0]
            if RootSliceWordDocumentChunk[1] != 0:
                try:
                    thePosition = (int(ReverseHex(Header2FourBytes(RootSliceWordDocumentChunk[1])[29]),16)+1)*512
                    if thePosition == 512:
                        RootSliceWordDocumentPosition = 0
                    else:
                        RootSliceWordDocumentPosition = thePosition
                except ValueError:
                    RootSliceWordDocumentPosition = 0
                try:
                    RootSliceWordDocumentLength = int(ReverseHex(Header2FourBytes(RootSliceWordDocumentChunk[1])[30]),16)
                except ValueError:
                    RootSliceWordDocumentLength =0
            else:
                RootSliceWordDocumentPosition = 0
                RootSliceWordDocumentLength = 0
        
            # 6) SummaryInformation
            
            if "05005300" in alternateRootSlice:
                try:
                    RootSliceSummaryInformationChunk = PositionChunkList(alternateRootSlice, "05005300", 256)
                except ValueError:
                    pass
            else:
                RootSliceSummaryInformationChunk = [0,0]
            if RootSliceSummaryInformationChunk[1] != 0:
                try:
                    thePosition = (int(ReverseHex(Header2FourBytes(RootSliceSummaryInformationChunk[1])[29]),16)+1)*512
                    if thePosition == 512:
                        RootSliceSummaryInformationPosition = 0
                    else:
                        RootSliceSummaryInformationPosition = thePosition
                except ValueError:
                    RootSliceSummaryInformationPosition = 0
                try:
                    RootSliceSummaryInformationLength = int(ReverseHex(Header2FourBytes(RootSliceSummaryInformationChunk[1])[30]),16)
                except ValueError:
                    RootSliceSummaryInformationLength =0
            else:
                RootSliceSummaryInformationPosition = 0
                RootSliceSummaryInformationLength = 0
        
            # 7) CompObject
    
            if "01004300" in alternateRootSlice:
                try:
                    RootSliceCompObjectChunk = PositionChunkList(alternateRootSlice, "01004300", 256)
                except ValueError:
                    pass
            else:
                RootSliceCompObjectChunk = [0,0]
            if RootSliceCompObjectChunk[1] != 0:
                try:
                    thePosition = (int(ReverseHex(Header2FourBytes(RootSliceCompObjectChunk[1])[29]),16)+1)*512
                    if thePosition == 512:
                        RootSliceCompObjectPosition = 0
                    else:
                        RootSliceCompObjectPosition = thePosition
                except ValueError:
                    RootSliceCompObjectPosition = 0
                try:
                    RootSliceCompObjectLength = int(ReverseHex(Header2FourBytes(RootSliceCompObjectChunk[1])[30]),16)
                except ValueError:
                    RootSliceCompObjectLength =0
            else:
                RootSliceCompObjectPosition = 0
                RootSliceCompObjectLength = 0
        
            # 8) Workbook
    
            if "57006f0072006b00" in alternateRootSlice:
                try:
                    RootSliceWorkbookChunk = PositionChunkList(alternateRootSlice, "57006f0072006b00", 256)
                except ValueError:
                    pass
            else:
                RootSliceWorkbookChunk = [0,0]
            if RootSliceWorkbookChunk[1] != 0:
                try:
                    thePosition = (int(ReverseHex(Header2FourBytes(RootSliceWorkbookChunk[1])[29]),16)+1)*512
                    if thePosition == 512:
                        RootSliceWorkbookPosition = 0
                    else:
                        RootSliceWorkbookPosition = thePosition
                except ValueError:
                    RootSliceWorkbookPosition = 0
                try:
                    RootSliceWorkbookLength = int(ReverseHex(Header2FourBytes(RootSliceWorkbookChunk[1])[30]),16)
                except ValueError:
                    RootSliceWorkbookLength =0
            else:
                RootSliceWorkbookPosition = 0
                RootSliceWorkbookLength = 0
        
            # 9) Sheet
    
            if "53006800" in alternateRootSlice:
                try:
                    RootSliceSheetChunk = PositionChunkList(alternateRootSlice, "53006800", 256)
                except ValueError:
                    pass
            else:
                RootSliceSheetChunk = [0,0]
            if RootSliceSheetChunk[1] != 0:
                try:
                    thePosition = (int(ReverseHex(Header2FourBytes(RootSliceSheetChunk[1])[29]),16)+1)*512
                    if thePosition == 512:
                        RootSliceSheetPosition = 0
                    else:
                        RootSliceSheetPosition = thePosition
                except ValueError:
                    RootSliceSheetPosition = 0
                try:
                    RootSliceSheetLength = int(ReverseHex(Header2FourBytes(RootSliceSheetChunk[1])[30]),16)
                except ValueError:
                    RootSliceSheetLength =0
            else:
                RootSliceSheetPosition = 0
                RootSliceSheetLength = 0
        
        
    DocSumList = []
    DocSumList.append("ALTERNATEDocumentSummary")
    DocSumList.append(RootSliceDocumentSummaryInfoChunk)
    DocSumList.append(RootDocSumInfoPosition)
    DocSumList.append(RootDocSumInfoPositionLength)
    if RootDocSumInfoPosition == 0:
        DocSumList.append(0)
    else:
        DocSumList.append(herx[2 * RootDocSumInfoPosition:2 * RootDocSumInfoPosition + RootDocSumInfoPositionLength])
    PositionsList.append(DocSumList)
    DataList=[]
    DataList.append("ALTERNATEData")
    DataList.append(RootSliceDataChunk)
    DataList.append(RootSliceDataPosition)
    DataList.append(RootSliceDataLength)
    if RootSliceDataPosition == 0:
        DataList.append(0)
    else:
        DataList.append(herx[2 * RootSliceDataPosition:2 * RootSliceDataPosition + RootSliceDataPositionLength])
    PositionsList.append(DataList)
    TableList=[]
    TableList.append("ALTERNATE1Table")
    TableList.append(RootSlice1TableChunk)
    TableList.append(RootSlice1TablePosition)
    TableList.append(RootSlice1TableLength)
    if RootSlice1TablePosition == 0:
        TableList.append(0)
    else:
        TableList.append(herx[2 * RootSlice1TablePosition:2 * RootSlice1TablePosition + RootSlice1TablePositionLength])
  
    PositionsList.append(TableList)
    WordDocumentList = []
    WordDocumentList.append("ALTERNATEWordDocument")
    WordDocumentList.append(RootSliceWordDocumentChunk)
    WordDocumentList.append(RootSliceWordDocumentPosition)
    WordDocumentList.append(RootSliceWordDocumentLength)
    if RootSliceWordDocumentPosition == 0:
        WordDocumentList.append(0)
    else:
        WordDocumentList.append(herx[2 * RootSliceWordDocumentPosition:2 * RootSliceWordDocumentPosition + RootSliceWordDocumentLength])
    PositionsList.append(WordDocumentList)
    SummaryInformationList = []
    SummaryInformationList.append("ALTERNATESummaryInformation")
    SummaryInformationList.append(RootSliceSummaryInformationChunk)
    SummaryInformationList.append(RootSliceSummaryInformationPosition)
    SummaryInformationList.append(RootSliceSummaryInformationLength)
    if RootSliceSummaryInformationPosition == 0:
        SummaryInformationList.append(0)
    else:
        SummaryInformationList.append(herx[2 * RootSliceSummaryInformationPosition:2 * RootSliceSummaryInformationPosition + RootSliceSummaryInformationLength])
    PositionsList.append(SummaryInformationList)
    CompObjectList = []
    CompObjectList.append("ALTERNATECompObject")
    CompObjectList.append(RootSliceCompObjectChunk)
    CompObjectList.append(RootSliceCompObjectPosition)
    CompObjectList.append(RootSliceCompObjectLength)
    if RootSliceCompObjectPosition == 0:
        CompObjectList.append(0)
    else:
        CompObjectList.append(herx[2 * RootSliceCompObjectPosition:2 * RootSliceCompObjectPosition + RootSliceCompObjectLength])
    PositionsList.append(CompObjectList)
    WorkbookList = []
    WorkbookList.append("ALTERNATEWorkbook")
    WorkbookList.append(RootSliceWorkbookChunk)
    WorkbookList.append(RootSliceWorkbookPosition)
    WorkbookList.append(RootSliceWorkbookLength)
    if RootSliceWorkbookPosition == 0:
        WorkbookList.append(0)
    else:
        WorkbookList.append(herx[2 * RootSliceWorkbookPosition:2 * RootSliceWorkbookPosition + RootSliceWorkbookLength])
    PositionsList.append(WorkbookList)
    SheetList = []
    SheetList.append("ALTERNATESheet")
    SheetList.append(RootSliceSheetChunk)
    SheetList.append(RootSliceSheetPosition)
    SheetList.append(RootSliceSheetLength)
    if RootSliceSheetPosition == 0:
        SheetList.append(0)
    else:
        SheetList.append(herx[2 * RootSliceSheetPosition:2 * RootSliceSheetPosition + RootSliceSheetLength])
    PositionsList.append(SheetList)

    return PositionsList
