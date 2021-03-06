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
    if siscontents[7] == "02d5cdd5":
        SISHeadDict["SISInfoGUID"] = "D5CDD502-2E9C-101B-9397-08002B2CF9AE"
    elif siscontents[7] == "05d5cdd5":
        SISHeadDict["SISUserGUID"] = "D5CDD505-2E9C-101B-9397-08002B2CF9AE"
    else:
        SISHeadDict["SISInfoGUID"] = siscontents[7]+siscontents[8]+siscontents[9]+siscontents[10]

    SISHeadDict["SISOffset"]=siscontents[11]
    sz = siscontents[12]
    SISHeadDict["SISLengthDec"] = int(sz[6:8]+sz[4:6]+sz[2:4]+sz[0:2],16)
    num = siscontents[13]
    SISHeadDict["NumberPropertiesDec"] = int(num[6:8]+num[4:6]+num[2:4]+num[0:2],16)
    return SISHeadDict

def suminfoParse(siscontents, num):

    numRes = siscontents[num]
    npd = int(numRes[6:8]+numRes[4:6]+numRes[2:4]+numRes[0:2],16)
    sisproplist = []
    x = 0
    while x < npd*2:
        indproplist = []
        indproplist.append(ReverseHex(siscontents[(num+1)+x]))
        indproplist.append(int(ReverseHex(siscontents[(num+1)+x+1]),16)+int(ReverseHex(siscontents[11]),16))
        x = x+2
        sisproplist.append(indproplist)
    return sisproplist

def SISTwo(siscontents):
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
    propcount =2
    propidentifieroffsetlist = []
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
        propidentifieroffsetlist.append(apropentry)
        propcount = propcount +2
    #########################
    # ADD FAKE ENTRY FOR END OF DICT, SO PARSING IS EASIER
    ######################################
    fakeend = []
    fakeend.append("End of Dict Chunk")
    fakeend.append(positionoffirstdictchunk+(UDDSDict["LenDictDec"]/4))
    propidentifieroffsetlist.append(fakeend)
    dickanswers = []
    userDick = {}
    for d in range(len(propidentifieroffsetlist)-1):
        thekey = propidentifieroffsetlist[d][0]
        startpos = positionoffirstdictchunk+propidentifieroffsetlist[d][1]
        endpos = positionoffirstdictchunk+propidentifieroffsetlist[d+1][1]
        finalpos = positionoffirstdictchunk+(UDDSDict["LenDictDec"]/4)

        if thekey == "Dictionary Property":
            
            try:
                UDDSDict["dicktophonedict"] = dicktophone(siscontents[startpos:endpos])
            except:
                print "is there problem with dictionary property"
                print siscontents[startpos:endpos]
        elif thekey == "CodePage Property":
            try:
                userDick["UserDefinedCodepage"] = int(ReverseHex(siscontents[startpos:endpos][-1]),16)
            except:
                print "problem codepage"
                print siscontents[startpos:endpos]
        elif thekey == "Locale Property":
            try:
                hislocale = checkLocale(str(siscontents[startpos:endpos][-1]))
                userDick["UserDefLocale"] = hislocale
            except:
                print "problem locale"
                print siscontents[startpos:endpos]
        elif thekey == "Behaviour Property":
            try:
                userDick["UserDefBehaviour"] = siscontents[startpos:endpos][-1]
            except:
                print "problem Behaviour"
                print siscontents[startpos:endpos]
        elif "Property Number" in thekey and "1f" in siscontents[startpos] :
            try:
                dickitem = []
                dickitem.append(thekey)
                hexdick = hextoascii(siscontents[startpos:endpos][2:])
                dickitem.append(hexdick)
                dickanswers.append(dickitem)
            except ValueError:
                try:
                    dickitem = []
                    dickitem.append(thekey)
                    hexdick = hextoascii(siscontents[startpos:finalpos][2:])
                    dickitem.append(hexdick)
                    dickanswers.append(dickitem)
                except:
                    #may not be ascii
                    dickitem = []
                    dickitem.append(thekey)
                    hexdick = siscontents[startpos:endpos][2:]
                    dickitem.append(hexdick)
                    dickanswers.append(dickitem)
        elif "Property Number" in thekey and "03000000" or "02000000" in siscontents[startpos]:
            # "03000000" is for a 32 bit integer 02 for 16 bit, may need to add in the others... 
            try:
                dickitem = []
                dickitem.append(thekey)
                intdick = int(ReverseHex(siscontents[startpos:endpos][1]),16)
                dickitem.append(intdick)
                dickanswers.append(dickitem)
            except:
                #may not be integer
                dickitem = []
                dickitem.append(thekey)
                intdick = siscontents[startpos:endpos][2:]
                dickitem.append(intdick)
                dickanswers.append(dickitem)
        else: 
            try:
                UDDSDict[thekey] = siscontents[startpos:endpos]
            except ValueError:
                UDDSDict[thekey] = siscontents[startpos:finalpos]


    for d in dickanswers:

        numberdick = int(d[0].split(":0")[1].split("0")[0])
        dickkey = UDDSDict["dicktophonedict"][numberdick]
        dickans = d[1]
        userDick[dickkey] = dickans
    UDDSDict["UserDictionary"] = userDick
    UDDSDict.pop("dicktophonedict", None)
    return UDDSDict


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

def dsmain(rawstuff):
    SISPropDict = {}
    siscontents =  Header2FourBytes(rawstuff)
    SISHead = SummaryInfoHead(siscontents)
    # SISHead returns correct dict of items
    SISPropDict["Document Summary Header"] = SISHead
    ng = numGuids(siscontents)
    SISPropDict["DWORD-cSections"] = ng
    if ng == 1:
        SISPropList = suminfoParse(siscontents, 13)
        finalsisproplist = sispropitems(SISPropList, siscontents)
        return finalsisproplist
    if ng == 2:
        SISPropList = suminfoParse(siscontents, 18)
        finalsisproplist = sispropitems(SISPropList, siscontents)
        UserDefPropDict = SISTwo(siscontents)
        SISPropDict["User Defined Section"] = UserDefPropDict
        return SISPropDict, finalsisproplist
    

