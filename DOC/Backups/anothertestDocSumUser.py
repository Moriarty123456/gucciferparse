#!/usr/bin/env python

import sys


sys.path.append("/home/david/Documents/ClintonJail/Guccifer2.0/Docs&Tools/Python/ALL_GUC_FILE_PROJECT/python/modules/DOC")
from binarytoolkit import *
from USERDEFDOCSUM import *
#from USERDEFDOCSUM import UDDSHead


rawchunk = "feff000005000200000000000000000000000000000000000200000002d5cdd59c2e1b10939708002b2cf9ae4400000005d5cdd59c2e1b10939708002b2cf9ae68010000240100000900000001000000500000000f0000005800000017000000680000000b000000700000001000000078000000130000008000000016000000880000000d000000900000000c000000e300000002000000e40400001e00000008000000444343430000000003000000a8190b000b000000000000000b000000000000000b000000000000000b000000000000001e1000000300000014000000546f702052656e6577616c2054617267657473000a00000050726f737065637473002100000027546f702052656e6577616c205461726765747327215072696e745f41726561000c100000040000001e0000000b000000576f726b7368656574730003000000020000001e0000000d0000004e616d65642052616e676573000300000001000000007c0500000a00000000000000580000000100000034010000020000003c010000030000000c05000004000000140500000500000020050000060000003405000007000000500500000800000068050000090000007005000008000000020000000c0000005f5049445f484c494e4b530003000000140000005f4164486f635265766965774379636c6549440004000000100000005f4e65775265766965774379636c6500050000000e0000005f456d61696c5375626a65637400060000000d0000005f417574686f72456d61696c0007000000180000005f417574686f72456d61696c446973706c61794e616d6500080000001c0000005f50726576696f75734164486f635265766965774379636c6549440009000000190000005f526576696577696e67546f6f6c7353686f776e4f6e63650002000000e404000041000000c80300003600000003000000070037000300000008000100030000000000000003000000060000001f0000001c0000006d00610069006c0074006f003a007000610074006800430061006e00400063006900740065006300680063006f002e006e006500740000001f00000001000000000000000300000009002b000300000007000100030000000000000003000000060000001f000000180000006d00610069006c0074006f003a006c00620065006c00640065006e00400068006800630063002e0063006f006d0000001f0000000100000000000000030000006b0040000300000006000100030000000000000003000000060000001f0000001a0000006d00610069006c0074006f003a006a0061007a007200610063006b004000730061006b006300610070002e0063006f006d0000001f0000000100000000000000030000000a0071000300000005000100030000000000000003000000060000001f0000001c0000006d00610069006c0074006f003a0063006b00610070006c0061006e00400067006c002d006e0079006c00610077002e0063006f006d0000001f000000010000000000000003000000760042000300000004000100030000000000000003000000060000001f0000001a0000006d00610069006c0074006f003a004300680072006900730040006d0061007300730032003000320030002e006f007200670000001f000000010000000000000003000000460077000300000003000100030000000000000003000000060000001f0000001a0000006d00610069006c0074006f003a006a0066006c0065006d0069006e00670040006f007800620069006f002e0063006f006d0000001f0000000100000000000000030000007a0053000300000002000100030000000000000003000000060000001f0000001f0000006d00610069006c0074006f003a00760065006c006d0061006c006500680040007700770068006f006c00640069006e00670073002e0063006f006d00000000001f000000010000000000000003000000460068000300000001000100030000000000000003000000060000001f0000001a0000006d00610069006c0074006f003a00680061006c006c0065006e00400061006c006c0065006e0063006f002e0063006f006d0000001f0000000100000000000000030000004c0072000300000000000100030000000000000003000000060000001f0000001f0000006d00610069006c0074006f003a00410072006f006e0073006f006e00400061006a006f0070006100720074006e006500720073002e0063006f006d00000000001f000000010000000000000003000000177e75c21e00000004000000000000001e0000000c00000052656e6577616c73000000001e000000140000004b6e7574736f6e40444343432e4f5247000000001e000000100000004861726d6f6e79204b6e7574736f6e0003000000b9e4f3331e00000004000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000"




def getpropidentifierdict(siscontents, positionoffirstdictchunk, numberproperties):
    propidentifieroffsetdict = {}
    ld = siscontents[positionoffirstdictchunk]
    lod = int(ld[6:8]+ld[4:6]+ld[2:4]+ld[0:2],16)
    numberproperties = int(ReverseHex(siscontents[positionoffirstdictchunk+1]),16)
    propcount =2
    for x in range(numberproperties):
        apropentry = []
        ap = siscontents[positionoffirstdictchunk+propcount]
        if str(ap) == "00000000":
            apropentry.append("Dictionary Property")
        elif str(ap) == "01000000":
            apropentry.append("CodePage Property")
        elif str(ap) == "00000080":
            apropentry.append("Locale Property")
        elif str(ap) == "03000080":
            apropentry.append("Behaviour Property")
        else:
            name = "Property Number :" + str(ap)
            apropentry.append(name)
        prof = siscontents[positionoffirstdictchunk+propcount+1]
        decprof = int(prof[6:8]+prof[4:6]+prof[2:4]+prof[0:2],16)
        chunkprof = decprof/4
        apropentry.append(chunkprof)
        propidentifieroffsetdict[apropentry[0]]=apropentry[1]
        propcount = propcount +2
            #########################
            # ADD FAKE ENTRY FOR END OF DICT, SO PARSING IS EASIER
            ######################################
    propidentifieroffsetdict["End of Dict Chunk"]=positionoffirstdictchunk+(lod/4)
    keyposition = []
    dickanswers = []
    userDick = {}
    for thekey in propidentifieroffsetdict:
        name = thekey
        position = propidentifieroffsetdict[thekey]
        keyposition.append(position)
        sortedpositions = keyposition.sort(cmp)
        #print keyposition -> [22, 77, 79, 323, 325, 328, 333, 340, 346, 348, 441]
    for k in range(len(keyposition)-1):
        itemposition = []
        itemposition.append(keyposition[k])
        itemposition.append(keyposition[k+1])
        for kk in propidentifieroffsetdict:
            if propidentifieroffsetdict[kk] == keyposition[k]:
                propidentifieroffsetdict[kk] = itemposition
        #propidentifieroffsetdict:
        #{'Property Number :07000000': [340, 346], 'Property Number :08000000': [346, 348], 'Property Number :03000000': [323, 325], 'Property Number :09000000': [348, 441], 'Property Number :04000000': [325, 328], 'CodePage Property': [77, 79], 'Property Number :05000000': [328, 333], 'End of Dict Chunk': 441, 'Property Number :06000000': [333, 340], 'Dictionary Property': [22, 77], 'Property Number :02000000': [79, 323]}

    return propidentifieroffsetdict


def UDDSTwo(siscontents):
    UDDSDict ={}
    # first offset is at 11
    fo = siscontents[11]
    UDDSDict["FirstOffsetDec"]=int(fo[6:8]+fo[4:6]+fo[2:4]+fo[0:2],16)
    UDDSDict["FirstOffsetInChunks"]= (UDDSDict["FirstOffsetDec"])/4
    # SecondOffset
    so = siscontents[16]
    UDDSDict["SecondOffsetDec"] = ((int(so[6:8]+so[4:6]+so[2:4]+so[0:2],16))/4)
    UDDSDict["SecondOffsetInChunks"] = (UDDSDict["SecondOffsetDec"])/4
    dslen = siscontents[17]
    docsumlen =  int(dslen[6:8]+dslen[4:6]+dslen[2:4]+dslen[0:2],16)
    UDDSDict["dslengthChunks"] =docsumlen/4 
    # UDDSDict["dslengthChunks"] gives the length of the doc summary part - the first part
    # adding it to the first offset gives the chunk at the start of the dict, which will be the length of the dict
    positionoffirstdictchunk = UDDSDict["FirstOffsetInChunks"] + UDDSDict["dslengthChunks"]
    lod = siscontents[positionoffirstdictchunk]
    UDDSDict["LenDictDec"] = int(lod[6:8]+lod[4:6]+lod[2:4]+lod[0:2],16)
    numberproperties = int(ReverseHex(siscontents[positionoffirstdictchunk+1]),16)
    propidentifieroffsetdict = getpropidentifierdict(siscontents, positionoffirstdictchunk, numberproperties)
    propidentifieroffsetdict.pop("End of Dict Chunk")
    dickanswers = []
    propertydict = {}
    for d in propidentifieroffsetdict:
        startpos = positionoffirstdictchunk+propidentifieroffsetdict[d][0]
        endpos = positionoffirstdictchunk+propidentifieroffsetdict[d][1]
        if d == "Dictionary Property":    
            UDDSDict["dicktophonedict"] = dicktophone(siscontents[startpos:endpos])
        elif "Property Number" in d:
            figure = int(ReverseHex(d.split(":")[1]),16)
            propertydict[figure] = siscontents[startpos:endpos]
        elif d == "CodePage Property":
            UDDSDict["UserDefinedCodepage"] = int(ReverseHex(siscontents[startpos:endpos][-1]),16)
        elif d == "Behaviour Property":
            UDDSDict["Behaviour Property"] = siscontents[startpos:endpos]
        elif d == "Locale Property":
            UDDSDict["UserDef Locale"] = checkLocale(str(siscontents[startpos:endpos][-1]))
        else:
            pass
    return UDDSDict, propertydict

def matchup(UserDefPropDict, propertydict):
    matchedproperties = []
    for key in UserDefPropDict["dicktophonedict"]:
        aline = []
        # print key - > 3
        # print UserDefPropDict["dicktophonedict"][key] - > _AdHocReviewCycleID
        # print propertydict[key] - > ['03000000', '177e75c2']
        aline.append(key)
        aline.append(UserDefPropDict["dicktophonedict"][key])
        aline.append(propertydict[key])
        matchedproperties.append(aline)
    return matchedproperties

def VT_Values(matchedproperties):
    VTMatchedDict = {}
    for v in matchedproperties:
        VT = v[2][0]
        LE = v[2][1]
        aVTentry=[]
        if VT == "03000000":
            aVTentry.append("VT: "+VT)
            aVTentry.append(int(ReverseText(LE),16))
            VTMatchedDict[v[1]]=aVTentry
        elif VT =="1e000000":
            aVTentry.append("VT: "+VT)
            aVTentry.append(hextoascii(v[2][2:]))
            VTMatchedDict[v[1]]=aVTentry
        elif VT =="41000000":
            if v[1] == "_PID_HLINKS":
                aVTentry.append("VT: "+VT)
                emailvector = PIDHlinks(v[2])
                aVTentry.append(emailvector)
                VTMatchedDict[v[1]]=aVTentry
            else:
                aVTentry.append("VT: "+VT)
                aVTentry.append(hextoascii(v[2][1:]))
                VTMatchedDict[v[1]]=aVTentry
        else:
            # print "entry is ", v[1]
            # print VT, "is unknown:", v[2][1:]
            aVTentry.append("VT_unknown: "+VT)
            aVTentry.append(v[2][1:])
            VTMatchedDict[v[1]]=aVTentry
    return VTMatchedDict

def popuseless(UserDefPropDict):
    
    poplist = ["dslengthChunks","FirstOffsetDec","SecondOffsetInChunks","SecondOffsetDec","LenDictDec","FirstOffsetInChunks"]
    for x in poplist:
        UserDefPropDict.pop(x,None)
    return UserDefPropDict

def newuddsmain(rawstuff):

    # INPUT: The Chunk containing the usder def doc sum bit.
    siscontents =  Header2FourBytes(rawstuff)
    UserDefPropDict, propertydict = UDDSTwo(siscontents)
    matchedproperties = matchup(UserDefPropDict, propertydict) #list
    VTMatchedDict = VT_Values(matchedproperties)
    
    UserDefPropDict["dicktophonedict"] = VTMatchedDict
    CleanUserDefPropDict = popuseless(UserDefPropDict)
    return CleanUserDefPropDict

#print newuddsmain(rawchunk)
# OUTPUT:
# {'UserDefinedCodepage': 1252, 'dicktophonedict': {'_AuthorEmail': ['VT: 1e000000', 'Knutson@DCCC.ORG'], '_AuthorEmailDisplayName': ['VT: 1e000000', 'Harmony Knutson'], '_AdHocReviewCycleID': ['VT: 03000000', 3262479895L], '_PID_HLINKS': ['VT: 41000000', ['mailto:pathCan@citechco.net', 'mailto:lbelden@hhcc.com', 'mailto:jazrack@sakcap.com', 'mailto:ckaplan@gl-nylaw.com', 'mailto:Chris@mass2020.org', 'mailto:jfleming@oxbio.com', 'mailto:velmaleh@wwholdings.com', 'mailto:hallen@allenco.com', 'mailto:Aronson@ajopartners.com']], '_EmailSubject': ['VT: 1e000000', 'Renewals'], '_NewReviewCycle': ['VT: 1e000000', ''], '_ReviewingToolsShownOnce': ['VT: 1e000000', ''], '_PreviousAdHocReviewCycleID': ['VT: 03000000', 871621817]}}
