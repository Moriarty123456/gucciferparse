import sys

sys.path.append("/home/david/Documents/ClintonJail/Guccifer2.0/Docs&Tools/Python/ALL_GUC_FILE_PROJECT/python/modules/DOC")

from binarytoolkit import *

def rootPARSE(infile):
    ###################################################
    # INPUT: a file path to a binary doc or xls file  #
    # OUTPUT: a dictionary of root attributes         #
    ###################################################
    OffSetChaseList = []
    binarydict = {}
    try:
        herx = OutputHex(infile)
        header = OutputHeader(herx)
 #       print "rootPARSE: "#, header
 #       print ""
        headlist =  Header2FourBytes(header)
        dsllist = ListDSL(headlist)
#        print "dsllist: ", dsllist
        root = ListRootOffset(dsllist, herx)
        print "root: ", root
        print ""
        print "rootlen: ", len(root)
        rootsection, rootin4bytes = RootSectionList(root, herx)
        rtlist=RootTimes(rootin4bytes)
        startingsectorloc, streamsize = StartingSectorLocation(rootin4bytes)
        chaselist = []
        for r in root:
            
            chaselist.append(infile)
            chaselist.append(r)
            chaselist.append(hex(r))
            chaselist.append(headlist)
            chaselist.append(rootin4bytes)
            chaselist.append(rtlist)
            chaselist.append(startingsectorloc)
            chaselist.append(streamsize)
            chaselist.append(rootsection)
            
        OffSetChaseList.append(chaselist)  
    except:
        print "Computer says 'no'"
    
    for line in OffSetChaseList:
        binarydict["Filename"] = line[0]
        binarydict["RootOffset Decimal"] = line[1]
        binarydict["RootOffset Hex"] = line[2]
        binarydict["Directory Name"] = line[4][0:16]
        binarydict["Root Timestamp"] = line[5]
        binarydict["Dir Name Length"] = line[4][16][0:2]
        binarydict["Object Type"] = line[4][16][5:6]
        binarydict["Red Black"] = line[4][16][6:]
        binarydict["Left Sibling ID"] = line[4][17]
        binarydict["Right Sibling ID"] = line[4][18]
        binarydict["Child ID"] = line[4][19]
        binarydict["CLSID"] = line[4][20:24]
        binarydict["State Bits"] = line[4][24]
        binarydict["Creation Time"] = line[4][25:27]
        binarydict["Modified Time"] = line[4][27:29]
        binarydict["Next SecLoc Raw"] = line[4][29]
        binarydict["Next SecSize Raw"] = line[4][30]
        binarydict["Next SecLoc Decimal"] = line[6]
        binarydict["Next SecLoc Hex"] = hex(line[6])
        binarydict["Next SecSize Decimal"] = line[7]
        binarydict["Next SecSize Hex"] = hex(line[7])
        start= line[6]*2
        end = (line[6]+(128*4))*2
        #print hex(start/2)
        #print Header2FourBytes(herx[start:end])
        
    return binarydict, OffSetChaseList


infile = "/home/david/Documents/ClintonJail/Guccifer2.0/Docs&Tools/Python/ALL_GUC_FILE_PROJECT/downloaded_docs/xls/out-of-region-ne-donors.xls"

rootPARSE(infile)
