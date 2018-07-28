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
    return SISHeadDict


def SISTwo(siscontents):
    SISPropDict = {}

    #First UUID
    FirstUUID = siscontents[7] + siscontents[8] + siscontents[9] + siscontents[10]
    SISPropDict["FirstUUID"] = FirstUUID
    print "FirstUUID", FirstUUID

    # first offset is at 11
    fo = siscontents[11]
    SISPropDict["FirstOffsetDec"]=int(fo[6:8]+fo[4:6]+fo[2:4]+fo[0:2],16)
    SISPropDict["FirstOffsetInChunks"]= (SISPropDict["FirstOffsetDec"])/4
    print "first offset: raw, dec, in no Cunks", fo, SISPropDict["FirstOffsetDec"], SISPropDict["FirstOffsetInChunks"]

    # SecondUUID
    SecondUUID = siscontents[12] + siscontents[13] + siscontents[14] + siscontents[15]
    SISPropDict["SecondUUID"] = SecondUUID
    print "SecondUUID", SecondUUID
    
    # SecondOffset
   
    so = siscontents[16]
    SISPropDict["SecondOffsetDec"] = ((int(so[6:8]+so[4:6]+so[2:4]+so[0:2],16))/4)
    SISPropDict["SecondOffsetInChunks"] = (SISPropDict["SecondOffsetDec"])/4
    print "Second Offset: ", so, SISPropDict["SecondOffsetDec"], SISPropDict["SecondOffsetInChunks"]
    #second length (=finish doc summary) at 17
    dslen = siscontents[17]
    docsumlen =  int(dslen[6:8]+dslen[4:6]+dslen[2:4]+dslen[0:2],16)
    SISPropDict["dslengthChunks"] =docsumlen/4 

    # SISPropDict["dslengthChunks"] gives the length of the doc summary part - the first part
    # adding it to the first offset gives the chunk at the start of the dict, which will be the length of the dict
    positionoffirstdictchunk = SISPropDict["FirstOffsetInChunks"] + SISPropDict["dslengthChunks"]
    print "first dict chunk at position ", positionoffirstdictchunk 
    print "contents of above, gives length of dict ", siscontents[positionoffirstdictchunk]
    lod = siscontents[positionoffirstdictchunk]
    SISPropDict["LenDictDec"] = int(lod[6:8]+lod[4:6]+lod[2:4]+lod[0:2],16)
    print "which is dec and chunks", SISPropDict["LenDictDec"], SISPropDict["LenDictDec"]/4
    numberproperties = int(ReverseHex(siscontents[positionoffirstdictchunk+1]),16)
    print "-=-=--==--=-=-=--=-="
    print "-=-=--==--=-=-=--=-="
    print "-=-=--==--=-=-=--=-="
    print numberproperties, "number of dictionary property identifiers and their offsets"
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
        # THE BELOW WORKS, BUT MAKE IT ASCII AND MARRY UP DICIONARY WITH DICKTOPHONE

    for d in range(len(propidentifieroffsetlist)):
        print propidentifieroffsetlist[d][0], " is "
        try:
            print siscontents[positionoffirstdictchunk+propidentifieroffsetlist[d][1]:positionoffirstdictchunk+propidentifieroffsetlist[d+1][1]]
        except ValueError:
            print siscontents[positionoffirstdictchunk+propidentifieroffsetlist[d][1]:positionoffirstdictchunk+(SISPropDict["LenDictDec"]/4)]

    return SISPropDict, propidentifieroffsetlist

#WE ARE HERE!

#USE THE DICKTOPHONE FUNCTION

#USE THE CHECKLOCALE FUNCTION





def SummaryInfoSector(rawsector):
    # INPUT: The Chunk containing the usder def doc sum bit.
    SISList = []
    siscontents =  Header2FourBytes(rawsector)
    # siscontents returns correct section in form  ['feff0000', '05000200', '00000000', ...]
    SISHead = SummaryInfoHead(siscontents)
    # SISHead returns correct dict of items
    SISList.append(SISHead)
    # some document summaries have two guids. one for document summary and one for user defined info. The number of guids is given by DWORD-cSections.
    n = 2
    x = siscontents[6]
    p=int([x[i:i+n] for i in range(0, len(x), n)][0])
    # p returns the correct number of guids.
    SISList.append("DWORD-cSections:")
    SISList.append(p)
    SISPropDict, propidentifieroffsetlist = SISTwo(siscontents)


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

    sisdict, sispropertylist, siscontents = SummaryInfoSector(rawstuff)
    #finalsisproplist = sispropitems(sispropertylist, siscontents)
    #return sisdict, finalsisproplist
    return sisdict, sispropertylist
    

