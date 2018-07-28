import sys

sys.path.append("/home/david/Documents/ClintonJail/Guccifer2.0/Docs&Tools/Python/ALL_GUC_FILE_PROJECT/python/modules/DOC")

from binarytoolkit import *
from CompoundFileDirectoryEntry import *


    # The_List = [
    #     ("RootEntry" ,"52006f006f007400200045006e007400"),
    # ("DocumentSummary", "050044006f00630075006d0065006e00"),
    # ("Data", "44006100"),
    # ("1Table", "31005400610062006c00650000000000"),
    # ("WordDocument", "57006f007200640044006f0063007500"),
    # ("SummaryInformation", "0500530075006d006d00610072007900"),
    # ("CompObject", "010043006f006d0070004f0062006a00"),
    # ("Workbook", "57006f0072006b0062006f006f006b00"),
    # ("Sheet","53006800")
    # ]

    # try:
    #         RootSliceWorkbookChunk = PositionChunkList(officialRootSlice, "57006f0072006b00", 256)


def OfficialHeader(hexchunks):
    HEADERDict = {}
    MINIFATDict = {}
    DIFATDict = {}
    Header_Signature = "".join(hexchunks[0:2])
    HEADERDict["Header_Signature"] = Header_Signature
    if Header_Signature != "d0cf11e0a1b11ae1":
        isCFB = False
        HEADERDict["CFB Document?"] = isCFB
        return HEADERDict
    else:
        isCFB = True
        HEADERDict["CFB Document?"] = isCFB
        Header_CLSID = "".join(hexchunks[2:6])
        HEADERDict["Header_CLSID"] = Header_CLSID
        Minor_Version = hexchunks[6][:4]
        HEADERDict["Minor_Version"] = Minor_Version
        Major_Version = int(hexchunks[6][4:6],16)
        HEADERDict["Major_Version"] = Major_Version
        Byte_Order = hexchunks[7][:4]
        if Byte_Order == "feff":
            Endian = "Little Endian"
        elif Byte_Order == "fffe":
            Endian = "Big Endian"
        else:
            Endian = "Unknown or Error"
        HEADERDict["Byte Order"] = Endian
        Sector_Shift = hexchunks[7][4:]
        if Sector_Shift == "0900":
            secsize = "512Bytes = 200Hex"
        elif Sector_Shift == "0c00":
            secsize = "4096Bytes = 1000Hex"
        else:
            secsize = "Unknown or Error"
        HEADERDict["Sector Size"] = secsize
        Mini_Sector_Shift = ReverseInt(hexchunks[8])
        if Mini_Sector_Shift == 6:
            minsecsize = "64Bytes = 40Hex"
        else:
            minsecsize = "Unknown or Error"
        HEADERDict["Mini Sector Size"] = minsecsize
        Number_of_Directory_Sectors = hexchunks[10]
        if Number_of_Directory_Sectors == "00000000":
            nds = "Zero (if Major_Version 3 will be Zero)"
        else:
            nds = Number_of_Directory_Sectors
        HEADERDict["Number_of_Directory_Sectors"] = nds
        Number_of_FAT_Sectors = ReverseInt(hexchunks[11])
        HEADERDict["Number of FAT Sectors"] = Number_of_FAT_Sectors
        First_Directory_Sector_Location = (ReverseInt(hexchunks[12])+1)*512
        HEADERDict["First Directory Sector Location (Root)"] = First_Directory_Sector_Location
        Transaction_Signature_Number = ReverseInt(hexchunks[13])
        HEADERDict["Transaction Signature Number"] = Transaction_Signature_Number
        Minstream_Cutoff_Size = ReverseInt(hexchunks[14])
        if Minstream_Cutoff_Size == 4096:
            mcs = "4096Bytes or 1000Hex"
        else:
            mcs = "Unknown or Error"
        HEADERDict["Minstream Cutoff Size"] = mcs
        # MiniFATs
        First_MiniFat_Sector_Location = (ReverseInt(hexchunks[15])+1)*512
        MINIFATDict["First MiniFat Sector Location"] = First_MiniFat_Sector_Location
        Number_of_Mini_FAT_Sectors = ReverseInt(hexchunks[16])
        MINIFATDict["Number of Mini FAT Sectors"] = Number_of_Mini_FAT_Sectors
        HEADERDict["MINIFATDict"] = MINIFATDict
        First_DIFAT_Sector_Location = hexchunks[17]
        if First_DIFAT_Sector_Location == "feffffff":
            fdsl = "End Of Chain"
        else:
            fdsl = (ReverseInt(First_DIFAT_Sector_Location)+1)*512
        DIFATDict["First DIFAT Sector Location"] = fdsl
        Number_of_DIFAT_Sectors = ReverseInt(hexchunks[18])
        DIFATDict["Number of DIFAT Sectors"] = Number_of_DIFAT_Sectors
        DIFATList =[]
        for x in range(19,109):
            indDIFAT = []
            if hexchunks[x] == "ffffffff":
                pass
            else:
                indDIFAT.append(x-18)
                indDIFAT.append((ReverseInt(hexchunks[x])+1)*512)
                DIFATList.append(indDIFAT)
        DIFATDict["DIFATList"] = DIFATList
        HEADERDict["DIFATDict"] = DIFATDict
        
        return HEADERDict

def LastDIFAT(OFFICIALDict, hexchunks):
    lastDFList = []
    LastDIFATPosition = OFFICIALDict["HEADERINFO"]["DIFATDict"]["DIFATList"][-1][1]
    LDSectorStart = ReverseInt(hexchunks[LastDIFATPosition/4])
    lastDFList.append(LastDIFATPosition)
    lastDFList.append(LDSectorStart)
    return lastDFList
   

def OfficialPositions(infile):
    OFFICIALDict = {}
    filename = infile.split("/")[-1]
    OFFICIALDict["FILENAME"] = filename
    herx = OutputHex(infile)
    hexchunks = Header2FourBytes(herx)
    headdict = OfficialHeader(hexchunks)
    OFFICIALDict["HEADERINFO"] = headdict
    LocationofLastDIFATChanged = LastDIFAT(OFFICIALDict, hexchunks)
    OFFICIALDict["Last DIFAT Changed (Location & Sector)"] = LocationofLastDIFATChanged
    RootSlicePosition = OFFICIALDict["HEADERINFO"]["First Directory Sector Location (Root)"]
    RootSlice = herx[2*RootSlicePosition:(2*RootSlicePosition)+1024]
    OFFICIALDict["RootDirectoryDict"] = CompoundFileDirectoryEntry(RootSlice)
    return OFFICIALDict

def JustCFDLocations(CFD_Dict):
    CFDLocDict = {}
    for directory in CFD_Dict["RootDirectoryDict"]:
        loc_size_list = []
        loc_size_list.append(CFD_Dict["RootDirectoryDict"][directory]["Starting Sector Location"])
        loc_size_list.append(CFD_Dict["RootDirectoryDict"][directory]["Stream Size"])
        CFDLocDict[directory] = loc_size_list
    return CFDLocDict

def CFDLocationsINChunks(CFD_Dict):
    CFDLCDict = {}
    for directory in CFD_Dict["RootDirectoryDict"]:
        CFDLCDict[directory] ={}
        st_chu = (CFD_Dict["RootDirectoryDict"][directory]["Starting Sector Location"])/4
        le_chu = (CFD_Dict["RootDirectoryDict"][directory]["Stream Size"])/4
        fin_chu = st_chu+le_chu
        if st_chu == 0:
            CFDLCDict[directory]["Start"] = "In MINIFATDict"
            CFDLCDict[directory]["Length"] = le_chu            
        else:
            CFDLCDict[directory]["Start"] = st_chu
            CFDLCDict[directory]["End"] = fin_chu
    return CFDLCDict



infile = "/home/david/Documents/ClintonJail/Guccifer2.0/Docs&Tools/Python/ALL_GUC_FILE_PROJECT/downloaded_docs/xls/1st-quarter-targeted-renewals.xls"

result = OfficialPositions(infile)
CFDBytes = JustCFDLocations(result)
CFDChunks = CFDLocationsINChunks(result)
print CFDChunks
theDIFATlist= result["HEADERINFO"]["DIFATDict"]["DIFATList"]
theMINIFATlist= result["HEADERINFO"]["MINIFATDict"]


print "DIFAT"
for a,b in theDIFATlist:

    print a, b, hex((b-1)/512), hex(b)

print theMINIFATlist
