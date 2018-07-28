import sys

sys.path.append("/home/david/Documents/ClintonJail/Guccifer2.0/Docs&Tools/Python/ALL_GUC_FILE_PROJECT/python/modules/DOC")

from binarytoolkit import *

def rootPARSE(rootsection):
    ###################################################
    # INPUT: a file path to a binary doc or xls file  #
    # OUTPUT: a dictionary of root attributes         #
    ###################################################
    OffSetChaseList = []
    binarydict = {}
    rootin4bytes  = Header2FourBytes(rootsection)
    rtlist=RootTimes(rootin4bytes)
    startingsectorloc, streamsize = StartingSectorLocation(rootin4bytes)
    chaselist = []
    chaselist.append(rootin4bytes)
    chaselist.append(rtlist)
    chaselist.append(startingsectorloc)
    chaselist.append(streamsize)            
    OffSetChaseList.append(chaselist)  
    
    for line in OffSetChaseList:
        roughtext = "".join(line[0][0:16]).decode("hex")
        thetext = "".join(roughtext.split("\x00"))
        binarydict["Directory Name"] = thetext
        #binarydict["Root Timestamp"] = line[1]
        binarydict["Dir Name Length"] = line[0][16][0:2]
        binarydict["Object Type"] = line[0][16][5:6]
        binarydict["Red Black"] = line[0][16][6:]
        binarydict["Left Sibling ID"] = line[0][17]
        binarydict["Right Sibling ID"] = line[0][18]
        binarydict["Child ID"] = line[0][19]
        binarydict["CLSID"] = line[0][20:24]
        binarydict["State Bits"] = line[0][24]
        try:
            binarydict["Creation Time"] = TimeStamp(line[0][25:27])
        except ValueError:
            binarydict["Creation Time"] = "".join(line[0][25:27])
        try:
            binarydict["Modified Time"] = TimeStamp(line[0][27:29])
        except ValueError:
            binarydict["Modified Time"] = "".join(line[0][27:29])
        nextseclist = []
        if line[0][29] == "00000000" or line[0][29].lower() =="feffffff":
            nextseclist.append("00000000")
        else:
            rawlocation = line[0][29]
            thelocation = (ReverseInt(rawlocation)+1)*512
            nextseclist.append(thelocation)
        if line[0][30] == "00000000":
            nextseclist.append("00000000")
        else:
            rawsize = line[0][30]
            thelocationsize = ReverseInt(rawsize)
            nextseclist.append(thelocationsize)
        binarydict["Next Sector Location & Size"] = nextseclist
    return binarydict

