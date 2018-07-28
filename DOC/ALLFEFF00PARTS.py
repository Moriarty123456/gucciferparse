import sys

sys.path.append("/home/david/Documents/ClintonJail/Guccifer2.0/Docs&Tools/Python/ALL_GUC_FILE_PROJECT/python/modules/DOC")

from binarytoolkit import *


#Summary Info GUID = 'e0859ff2f94f6810ab9108002b27b3d9'
#DocSummary Info GUID ='02d5cdd59c2e1b10939708002b2cf9ae'
#UserDefn GUID = '05d5cdd59c2e1b10939708002b2cf9ae'
#Alternate GUID = '55c18c4c1e6cd1118e4100c04fb9386d'
#PropertyBag GUID = '01180020e65dd1118e3800c04fb9386d'

def FEFFPositions(infile):
    herx = OutputHex(infile)
    header = OutputHeader(herx)
    #WAHHH 2048 ISN'T LONG ENOUGH. SHOULD FIND THE END OF CHUNK'
    AllFEFFPositions, AllFEFFChunks = OutputFFslist(herx, "feff0000",4096)
    return AllFEFFPositions, AllFEFFChunks

def FEFFLists(infile):
    positions, chunks = FEFFPositions(infile)
    counter = 0
    FEFFSumInfoList = []
    FEFFDocSumUserDefList = []
    FEFFDocSumList = []
    FEFFUserDefList = []
    FEFFAlternateList = []
    FEFFPropertyBagList = []
    for chunk in chunks:
        if "e0859ff2f94f6810ab9108002b27b3d9" in chunk:
            FEFFSumInfoList.append(positions[counter])
            FEFFSumInfoList.append(chunks[counter])
        elif "02d5cdd59c2e1b10939708002b2cf9ae" and "05d5cdd59c2e1b10939708002b2cf9ae"in chunk:
            FEFFDocSumUserDefList.append(positions[counter])
            FEFFDocSumUserDefList.append(chunks[counter])
        elif "02d5cdd59c2e1b10939708002b2cf9ae" in chunk:
            FEFFDocSumList.append(positions[counter])
            FEFFDocSumList.append(chunks[counter])
        elif "05d5cdd59c2e1b10939708002b2cf9ae" in chunk:
            FEFFUserDefList.append(positions[counter])
            FEFFUserDefList.append(chunks[counter])
        elif "55c18c4c1e6cd1118e4100c04fb9386d" in chunk:
            FEFFAlternateList.append(positions[counter])
            FEFFAlternateList.append(chunks[counter])
        elif "01180020e65dd1118e3800c04fb9386d" in chunk:
            FEFFPropertyBagList.append(positions[counter])
            FEFFPropertyBagList.append(chunks[counter])
        else:
            print "NO Guid Matches!"
        counter = counter+1
    return FEFFSumInfoList, FEFFDocSumUserDefList, FEFFDocSumList, FEFFUserDefList, FEFFAlternateList, FEFFPropertyBagList

# infile = "/home/david/Documents/ClintonJail/Guccifer2.0/Docs&Tools/Python/ALL_GUC_FILE_PROJECT/downloaded_docs/xls/out-of-region-ne-donors.xls"

# A,B,C,D,E,F = FEFFLists(infile)

# print hex(A[0])
# print hex(B[0])
# print hex(C[0])
# print hex(D[0])
# print hex(E[0])
# print hex(F[0])
