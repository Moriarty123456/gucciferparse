import sys

sys.path.append("/home/david/Documents/ClintonJail/Guccifer2.0/Docs&Tools/Python/ALL_GUC_FILE_PROJECT/python/modules/DOC")

from binarytoolkit import *

def rawentrychunklist(rootchunks):
    #four entries of 32 chunks in one 128 chunk rootchunk. 
    countentry = 0
    rawentrylist = []
    for entry in range(1,5):
        startchunk = countentry
        endchunk = startchunk+31
        rawentrylist.append(rootchunks[startchunk:endchunk])
        countentry= countentry+32
    return rawentrylist


def CompoundFileDirectoryEntry(RootSlice):
    RootDirectoryDict = {}
    rootchunks = Header2FourBytes(RootSlice)
    listofrawentrychunks = rawentrychunklist(rootchunks)
    for entry in listofrawentrychunks:
        anentrydict = {}
        D_E_N = hextoascii("".join(entry[0:16]))
        if D_E_N[0] == '\x05' or  D_E_N[0] == '\x02' or  D_E_N[0] == '\x01':
            Directory_Entry_Name = D_E_N[1:]
        else:
            Directory_Entry_Name = D_E_N
        Directory_Entry_Name_Length =  ReverseInt(entry[16][:4])
        anentrydict["Directory Entry Name Length"] = Directory_Entry_Name_Length
        ot = entry[16][4:6]
        if ot == "05":
            OType = "Root Storage Object"
        elif ot == "02":
            OType = "Stream Object"
        elif ot == "01":
            OType = "Storage Object"
        elif ot == "00":
            OType = "Undefined Object"
        else:
            OType = "Unknown or Error"
        anentrydict["Object Type"] =  OType
        cf = entry[16][6:]
        if cf == "00":
            CFlag = "Red"
        elif cf == "01":
            CFlag = "Black"
        else:
            CFlag = "Unknown or Error"
        anentrydict["Color Flag"] = CFlag
        anentrydict["Left Sibling ID"] = entry[17]
        anentrydict["Right Sibling ID"] = entry[18]
        anentrydict["Child ID"] = entry[19]
        anentrydict["CLSID"] = "".join(entry[20:24])
        anentrydict["State Bits"] = entry[24]
        ts = entry[25:27]
        if ts[0] == "00000000" or ts[1] == "00000000":
            theCT = "None"
        else:
            theCT = TimeStamp(ts)
        anentrydict["Creation Time"] = theCT
        ms = entry[27:29]
        if ms[0] == "00000000" or ms[1] == "00000000":
            theMT = "None"
        else:
            theMT = TimeStamp(ms)
        anentrydict["Modified Time"] = theMT
        ssl = (ReverseInt(entry[29])+1)*512
        if ssl == 512:
            theSSL = 0
        else:
            theSSL = ssl
        # Note: if SSL = 0 then this is sector #0 in mini FAT.
        anentrydict["Starting Sector Location"] = theSSL 
        anentrydict["Stream Size"] = ReverseInt(entry[30])
        RootDirectoryDict[Directory_Entry_Name] = anentrydict
    return RootDirectoryDict
