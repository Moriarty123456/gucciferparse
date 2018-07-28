
import sys

sys.path.append("/home/david/Documents/ClintonJail/Guccifer2.0/Docs&Tools/Python/ALL_GUC_FILE_PROJECT/python/modules/DOC")

from binarytoolkit import *


def ReverseHex(somehex):
    return hex(int(somehex[6:8]+somehex[4:6]+somehex[2:4]+somehex[0:2],16))

def SummaryInfoHead(siscontents):
    SISHeadDict = {}
    SISHeadDict["WORDbyteOrder"]= siscontents[0][:4]
    SISHeadDict["WORD-version"] = siscontents[0][4:]
    SISHeadDict["OSMajorVersion"] = siscontents[1][:2]
    SISHeadDict["OSMinorVersion"] = siscontents[1][2:4]
    SISHeadDict["OSType"] = siscontents[1][4:8]
    SISHeadDict["GUID-applicationClsid"] = siscontents[2]+siscontents[3]+siscontents[4]+siscontents[5]
    #First UUID
    SISHeadDict["FirstUUID"] = siscontents[7] + siscontents[8] + siscontents[9] + siscontents[10]
    # SecondUUID
    SISHeadDict["SecondUUID"] = siscontents[12] + siscontents[13] + siscontents[14] + siscontents[15]
    SISHeadDict["NumberPropertiesDec"] = int(siscontents[18][6:8]+siscontents[18][4:6]+siscontents[18][2:4]+siscontents[18][0:2],16)
    return SISHeadDict

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


def suminfoParse(siscontents):
    num = siscontents[18]

    npd = int(num[6:8]+num[4:6]+num[2:4]+num[0:2],16)
    sisproplist = []
    x = 0
    while x < npd*2:
        indproplist = []
        indproplist.append(ReverseHex(siscontents[19+x]))
        indproplist.append(int(ReverseHex(siscontents[19+x+1]),16)+int(ReverseHex(siscontents[11]),16))
        x = x+2
        sisproplist.append(indproplist)
    return sisproplist
    
def sispropitems(sisproplist, siscontents):
    ############################################################
    # INPUT: the list of properties ("0x1", 184),              #
    # INPUT: the contents of the Summary info stream in chunks #
    # OUTPUT: Dictionary of meaning and value.                 #
    ############################################################
    theproperties = [("GKPIDDSI_CODEPAGE", "0x00000001", "0x1", "02000000", "no", "4"), ("GKPIDDSI_CATEGORY", "0x00000002", "0x2", "1e000000", "yes", "ANSI"), ("GKPIDDSI_PRESFORMAT", "0x00000003", "0x3", "1e000000", "yes", "ANSI"), ("GKPIDDSI_BYTECOUNT", "0x00000004", "0x4", "03000000", "no", "4"), ("GKPIDDSI_LINECOUNT", "0x00000005", "0x5", "03000000", "no", "4"), ("GKPIDDSI_PARACOUNT", "0x00000006", "0x6", "03000000", "no", "4"), ("GKPIDDSI_SLIDECOUNT", "0x00000007", "0x7", "03000000", "no", "4"), ("GKPIDDSI_NOTECOUNT", "0x00000008", "0x8", "03000000", "no", "4"), ("GKPIDDSI_HIDDENCOUNT", "0x00000009", "0x9", "03000000", "no", "4"), ("GKPIDDSI_MMCLIPCOUNT", "0x0000000A", "0xa", "03000000", "no", "4"), ("GKPIDDSI_SCALE", "0x0000000B", "0xb", "0b000000", "no", "4"), ("GKPIDDSI_HEADINGPAIR", "0x0000000C", "0xc", "000c1000", "no", "ANSI"), ("GKPIDDSI_DOCPARTS", "0x0000000D", "0xd", "1e100000", "yes", "ANSI"), ("GKPIDDSI_MANAGER", "0x0000000E", "0xe", "", "", ""), ("GKPIDDSI_COMPANY", "0x0000000F", "0xf", "1e000000", "yes", "ANSI"), ("GKPIDDSI_LINKSDIRTY", "0x00000010", "0x10", "0b000000", "no", "4"), ("GKPIDDSI_CCHWITHSPACES", "0x00000011", "0x11", "03000000", "no", "4"), ("GKPIDDSI_SHAREDDOC", "0x00000013", "0x13", "0b000000", "no", "4"), ("GKPIDDSI_LINKBASE", "0x00000014", "0x14", "41000000", "no", "ANSI"), ("GKPIDDSI_HLINKS", "0x00000015", "0x15", "41000000", "no", "ANSI"), ("GKPIDDSI_HYPERLINKSCHANGED", "0x00000016", "0x16", "0b000000", "no", "4"), ("GKPIDDSI_VERSION", "0x00000017", "0x17", "03000000", "no", "4"), ("GKPIDDSI_DIGSIG", "0x00000018", "0x18", "41000000", "no", "4"), ("GKPIDDSI_CONTENTTYPE", "0x0000001A", "0x1a", "1e000000", "yes", "ANSI"), ("GKPIDDSI_CONTENTSTATUS", "0x0000001B", "0x1b", "1e000000", "yes", "ANSI"), ("GKPIDDSI_LANGUAGE", "0x0000001C", "0x1c", "1e000000", "yes", "ANSI"), ("GKPIDDSI_DOCVERSION", "0x0000001D", "0x1d", "1e000000", "yes", "ANSI")]
    itemlist =[]
    chunklist = []
    for item in sisproplist:
        chunklist.append(item[1]/4)
        itemlist.append(item[0])
    lengths = []
    for x in range(len(chunklist)-1):
        l = chunklist[x+1]-chunklist[x]
        lengths.append(l)
    # to make up for difference in length of last item
    # which we know is 4 bytes or 1 chunk.
    lengths.append(1)
    finalsisproplist=[]
    for x in range(len(sisproplist)):
        itemcode = itemlist[x]
        for y in range(len(theproperties)):
            if itemcode == theproperties[y][2]:
                finalsisproplist.append(theproperties[y][0])
                lenmarker = theproperties[y][4]
                ansimarker = theproperties[y][5]
            else:
                pass
        start = chunklist[x]
        end = chunklist[x]+lengths[x]
        itemcont = siscontents[start:end]
        if lenmarker == "no":
            nottext = itemcont[1:]
            if len(nottext) == 1:
                notstring = "".join(nottext)
                revthefield = ReverseHex(notstring)
                thefield = int(revthefield,16)
                finalsisproplist.append(thefield)
            elif len(nottext) == 2 and itemcode == "0xa":
                #item will be a edittime
                finalsisproplist.append(int(EditTime(nottext)))
            elif len(nottext) == 2 and itemcode != "0xa":
                #item will be datetime
                finalsisproplist.append(TimeStamp(nottext))
            else:
                #"item is empty", nottext
                pass
        elif ansimarker == "ANSI":
            #should be text
            thetext = "".join(itemcont[2:]).decode("hex")
            if "\x00" in thetext:
                newtext = thetext.split("\x00")[0]
                finalsisproplist.append(newtext)
            else:
                finalsisproplist.append(thetext)
        else:
            pass
    return finalsisproplist



def uddsmain(rawstuff):

    # INPUT: The Chunk containing the usder def doc sum bit.
    SISPropDict = {}
    siscontents =  Header2FourBytes(rawstuff)
    # siscontents returns correct section in form  ['feff0000', '05000200', '00000000', ...]
    SISHead = SummaryInfoHead(siscontents)
    # SISHead returns correct dict of items
    SISPropDict["UserDefDocSum Header"] = SISHead
    # some document summaries have two guids. one for document summary and one for user defined info. The number of guids is given by DWORD-cSections.
    ng = numGuids(siscontents)
    SISPropDict["DWORD-cSections"] = ng
    # GET USER DEFINED SECTION
    UserDefPropDict, propertydict = UDDSTwo(siscontents)
    matchedproperties = matchup(UserDefPropDict, propertydict) #list
    VTMatchedDict = VT_Values(matchedproperties)
    UserDefPropDict["dicktophonedict"] = VTMatchedDict
    CleanUserDefPropDict = popuseless(UserDefPropDict)
    SISPropDict["User Defined Section"] = CleanUserDefPropDict
    # GET SUMMARY INFORMATION REGULAR SECTION
    sispropertylist = suminfoParse(siscontents)
    finalsisproplist = sispropitems(sispropertylist, siscontents)
    SISPropDict["Summary Info Section"] = finalsisproplist
    return SISPropDict
