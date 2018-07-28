import sys

sys.path.append("/home/david/Documents/ClintonJail/Guccifer2.0/Docs&Tools/Python/ALL_GUC_FILE_PROJECT/python/modules/DOC")

from binarytoolkit import *

def OfficialPositions(infile):
    PositionsList = []
    filename = infile.split("/")[-1]
    PositionsList.append("OFFICIAL "+filename)
    herx = OutputHex(infile)
    # get official location for root entry
    # 1) Get Header2FourBytes
    header = OutputHeader(herx)
    # 2) Get position, convert to dec.
    officialrootentryPosition = (int(ReverseHex(Header2FourBytes(header)[12]),16)+1)*512
    PositionsList.append(officialrootentryPosition)
    officialRootSlice = herx[2*officialrootentryPosition:(2*officialrootentryPosition)+1024]
        # Now, from the officialRootSlice string, get the official locations for docsum, docinfo, userdef, propbag, alternatestream if they exist (they may not).
    # If they don't exist then get chunks; name and locations, and lengths, as these may be things G2 added in ...
    # also get RootTime if it exists
    # 0) RootTimes
    headlist =  Header2FourBytes(header)
    dsllist = ListDSL(headlist)
    root = ListRootOffset(dsllist, herx)
    rootsection, rootin4bytes = RootSectionList(root, herx)
    rtlist=RootTimes(rootin4bytes)
    PositionsList.append("OFFICIALRootTimes")
    PositionsList.append(rtlist)
    PositionsList.append("OFFICIALRootSlice")
    PositionsList.append(officialRootSlice)    
    # 1) RootSliceDocumentSummaryInfo
      
    if "05004400" in officialRootSlice:
        try:
            RootSliceDocumentSummaryInfoChunk = PositionChunkList(officialRootSlice, "05004400", 256)
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
    
    if "44006100" in officialRootSlice:
        try:
            RootSliceDataChunk = PositionChunkList(officialRootSlice, "44006100", 256)
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

    if "31005400" in officialRootSlice:
        try:
            RootSlice1TableChunk = PositionChunkList(officialRootSlice, "31005400", 256)
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
    
    if "57006f0072006400" in officialRootSlice:
        try:
            RootSliceWordDocumentChunk = PositionChunkList(officialRootSlice, "57006f0072006400", 256)
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
    if "05005300" in officialRootSlice:
        try:
            RootSliceSummaryInformationChunk = PositionChunkList(officialRootSlice, "05005300", 256)
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
    
    if "01004300" in officialRootSlice:
        try:
            RootSliceCompObjectChunk = PositionChunkList(officialRootSlice, "01004300", 256)
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
    
    if "57006f0072006b00" in officialRootSlice:
        try:
            RootSliceWorkbookChunk = PositionChunkList(officialRootSlice, "57006f0072006b00", 256)
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
    
    if "53006800" in officialRootSlice:
        try:
            RootSliceSheetChunk = PositionChunkList(officialRootSlice, "53006800", 256)
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
    DocSumList.append("OFFICIALDocumentSummary")
    DocSumList.append(RootSliceDocumentSummaryInfoChunk)
    DocSumList.append(RootDocSumInfoPosition)
    DocSumList.append(RootDocSumInfoPositionLength)
    if RootDocSumInfoPosition == 0:
        DocSumList.append(0)
    else:
        DocSumList.append(herx[2 * RootDocSumInfoPosition:2 * RootDocSumInfoPosition + RootDocSumInfoPositionLength])
    PositionsList.append(DocSumList)
    DataList=[]
    DataList.append("OFFICIALData")
    DataList.append(RootSliceDataChunk)
    DataList.append(RootSliceDataPosition)
    DataList.append(RootSliceDataLength)
    if RootSliceDataPosition == 0:
        DataList.append(0)
    else:
        DataList.append(herx[2 * RootSliceDataPosition:2 * RootSliceDataPosition + RootSliceDataPositionLength])
    PositionsList.append(DataList)
    TableList=[]
    TableList.append("OFFICIAL1Table")
    TableList.append(RootSlice1TableChunk)
    TableList.append(RootSlice1TablePosition)
    TableList.append(RootSlice1TableLength)
    if RootSlice1TablePosition == 0:
        TableList.append(0)
    else:
        TableList.append(herx[2 * RootSlice1TablePosition:2 * RootSlice1TablePosition + RootSlice1TablePositionLength])
  
    PositionsList.append(TableList)
    WordDocumentList = []
    WordDocumentList.append("OFFICIALWordDocument")
    WordDocumentList.append(RootSliceWordDocumentChunk)
    WordDocumentList.append(RootSliceWordDocumentPosition)
    WordDocumentList.append(RootSliceWordDocumentLength)
    if RootSliceWordDocumentPosition == 0:
        WordDocumentList.append(0)
    else:
        WordDocumentList.append(herx[2 * RootSliceWordDocumentPosition:2 * RootSliceWordDocumentPosition + RootSliceWordDocumentLength])
    PositionsList.append(WordDocumentList)
    SummaryInformationList = []
    SummaryInformationList.append("OFFICIALSummaryInformation")
    SummaryInformationList.append(RootSliceSummaryInformationChunk)
    SummaryInformationList.append(RootSliceSummaryInformationPosition)
    SummaryInformationList.append(RootSliceSummaryInformationLength)
    if RootSliceSummaryInformationPosition == 0:
        SummaryInformationList.append(0)
    else:
        SummaryInformationList.append(herx[2 * RootSliceSummaryInformationPosition:2 * RootSliceSummaryInformationPosition + RootSliceSummaryInformationLength])
    PositionsList.append(SummaryInformationList)
    CompObjectList = []
    CompObjectList.append("OFFICIALCompObject")
    CompObjectList.append(RootSliceCompObjectChunk)
    CompObjectList.append(RootSliceCompObjectPosition)
    CompObjectList.append(RootSliceCompObjectLength)
    if RootSliceCompObjectPosition == 0:
        CompObjectList.append(0)
    else:
        CompObjectList.append(herx[2 * RootSliceCompObjectPosition:2 * RootSliceCompObjectPosition + RootSliceCompObjectLength])
    PositionsList.append(CompObjectList)
    WorkbookList = []
    WorkbookList.append("OFFICIALWorkbook")
    WorkbookList.append(RootSliceWorkbookChunk)
    WorkbookList.append(RootSliceWorkbookPosition)
    WorkbookList.append(RootSliceWorkbookLength)
    if RootSliceWorkbookPosition == 0:
        WorkbookList.append(0)
    else:
        WorkbookList.append(herx[2 * RootSliceWorkbookPosition:2 * RootSliceWorkbookPosition + RootSliceWorkbookLength])
    PositionsList.append(WorkbookList)
    SheetList = []
    SheetList.append("OFFICIALSheet")
    SheetList.append(RootSliceSheetChunk)
    SheetList.append(RootSliceSheetPosition)
    SheetList.append(RootSliceSheetLength)
    if RootSliceSheetPosition == 0:
        SheetList.append(0)
    else:
        SheetList.append(herx[2 * RootSliceSheetPosition:2 * RootSliceSheetPosition + RootSliceSheetLength])
    PositionsList.append(SheetList)
    
    return PositionsList

#infile = "/home/david/Documents/ClintonJail/Guccifer2.0/Docs&Tools/Python/ALL_GUC_FILE_PROJECT/downloaded_docs/xls/1st-quarter-targeted-renewals.xls"

#print OfficialPositions(infile)
# print len(stuff)
# #doc summary chunk given by:
# print stuff[7][4]
