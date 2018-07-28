import sys

sys.path.append("/home/david/Documents/ClintonJail/Guccifer2.0/Docs&Tools/Python/ALL_GUC_FILE_PROJECT/python/modules/DOC")

from binarytoolkit import *


def SummaryInfoLocation(infile):
    # Input: infile: the file name
    # Output: docInfoSectorList: a list of The chunks that 248 x 4 byte chunks that have the DSInfoLocations according to the root. But if one is blank it fails

    herx = OutputHex(infile)
    # some documents have two directory entries, so make list of directory entries.
    herxlist = herx.split("05004400")
    docInfoSectorList = []
    for x in herxlist:
        if x[0:10] == "6f00630075":
            sis ="05004400"+ x[0:248]
            #4ca00
            sislist =  Header2FourBytes(sis)
            # Location of Doc Sum Info is given by third to last chunk
            sl= sislist[-3]
            # check that this location is not blank:
            if int(sl,16) == 0:
                print "One DOCUMENT SUMMARY INFORMATION sector HAS A BLANK LOCATION"
            else:
                secloc = sl[6:8]+sl[4:6]+sl[2:4]+sl[0:2]
                # Size of Doc Sum Info is given by second to last chunk. Seems to usually be 00100000, or 4096 in dec.
                ss = sislist[-2]
                secsize = ss[6:8]+ss[4:6]+ss[2:4]+ss[0:2]
                # Calculate start and end of sector & append to list.
                decimalsectorlocation = 2*((int(secloc,16)+1)*512)
                endsidlocation = decimalsectorlocation + 2*(int(secsize,16))
                thechunk = herx[decimalsectorlocation:endsidlocation]
                docInfoSectorList.append(thechunk)
                #returns number of Summary Info Sectors, and the sector itself
    return docInfoSectorList


def ReverseHex(somehex):
    # In: some hex. Out the hex reversed
    return hex(int(somehex[6:8]+somehex[4:6]+somehex[2:4]+somehex[0:2],16))

def SummaryInfoHead(siscontents):
    # In: siscontents: the ds or sum info in 4 byte chunks
    # Out: the ostype, word version etc are added to SISHeadDict 

    SISHeadDict = {}
    SISHeadDict["WORDbyteOrder"]= siscontents[0][:4]
    SISHeadDict["WORD-version"] = siscontents[0][4:]
    SISHeadDict["OSMajorVersion"] = siscontents[1][:2]
    SISHeadDict["OSMinorVersion"] = siscontents[1][2:4]
    SISHeadDict["OSType"] = siscontents[1][4:8]
    SISHeadDict["GUID-applicationClsid"] = siscontents[2]+siscontents[3]+siscontents[4]+siscontents[5]
    return SISHeadDict

def SISOne(siscontents):
    SISPropDict = {}
    #we know the likely guids ...
    if siscontents[7] == "02d5cdd5":
        SISPropDict["SISInfoGUID"] = "{D5CDD502-2E9C-101B-9397-08002B2CF9AE}"
    elif siscontents[7] == "05d5cdd5":
        SISPropDict["SISUserGUID"] = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"
    else:
        SISPropDict["SISInfoGUID"] = siscontents[7]+siscontents[8]+siscontents[9]+siscontents[10]
    SISPropDict["SISOffset"]=siscontents[11]
    sl = siscontents[12]
    SISPropDict["SISLengthHex"] = hex(int(sl[6:8]+sl[4:6]+sl[2:4]+sl[0:2],16))
    sz = siscontents[12]
    SISPropDict["SISLengthDec"] = int(sz[6:8]+sz[4:6]+sz[2:4]+sz[0:2],16)
    num = siscontents[13]
    SISPropDict["NumberPropertiesHex"] = hex(int(num[6:8]+num[4:6]+num[2:4]+num[0:2],16))
    SISPropDict["NumberPropertiesDec"] = int(num[6:8]+num[4:6]+num[2:4]+num[0:2],16)
    npd = int(num[6:8]+num[4:6]+num[2:4]+num[0:2],16)
    print "docsuminfoPARSE Number of DocInfo Props: ", npd
    sisproplist = []
    x = 0
    while x < npd*2:
        indproplist = []
        print "docsuminfoPARSE line 48", siscontents[14+x]
        indproplist.append(ReverseHex(siscontents[14+x]))
        indproplist.append(int(ReverseHex(siscontents[14+x+1]),16)+int(ReverseHex(siscontents[11]),16))
        x = x+2
        sisproplist.append(indproplist)
    return sisproplist, SISPropDict

def SISTwo(siscontents):
    SISPropDict = {}
    #we know the likely guids ...
    if siscontents[12] == "05d5cdd5":
        SISPropDict["SISUserGUID"] = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"
    else:
        # there's something wrong. There should be the User Defined GUID Here. 
        print "there's something wrong. There should be the User Defined GUID Here."
        SISPropDict["SISInfoGUID"] = siscontents[12]+siscontents[13]+siscontents[14]+siscontents[15]
    so = siscontents[16]
    SISPropDict["SISOffsetHex"]=hex(int(so[6:8]+so[4:6]+so[2:4]+so[0:2],16))
    SISPropDict["SISOffsetDec"]=int(so[6:8]+so[4:6]+so[2:4]+so[0:2],16)

    #######################################
    # THIS IS THE START OF THE DICTIONARY #
    #######################################
    dictstart = siscontents[(int(so[6:8]+so[4:6]+so[2:4]+so[0:2],16))/4]
    print "dictstart: ", dictstart
    print "should be ac050000, for 1st quarter targeted.xls"

    ##########################################################
    # CONTINUE FROM HERE:                                    #
    # https://msdn.microsoft.com/en-us/library/dd942093.aspx #
    ##########################################################

    sl = siscontents[12]
    SISPropDict["SISLengthHex"] = hex(int(sl[6:8]+sl[4:6]+sl[2:4]+sl[0:2],16))
    sz = siscontents[12]
    SISPropDict["SISLengthDec"] = int(sz[6:8]+sz[4:6]+sz[2:4]+sz[0:2],16)
    num = siscontents[13]
    SISPropDict["NumberPropertiesHex"] = hex(int(num[6:8]+num[4:6]+num[2:4]+num[0:2],16))
    SISPropDict["NumberPropertiesDec"] = int(num[6:8]+num[4:6]+num[2:4]+num[0:2],16)
    npd = int(num[6:8]+num[4:6]+num[2:4]+num[0:2],16)
    print "docsuminfoPARSE Number of DocInfo Props: ", npd
    sisproplist = []
    x = 0
    while x < npd*2:
        indproplist = []
        print "docsuminfoPARSE line 48", siscontents[14+x]
        indproplist.append(ReverseHex(siscontents[14+x]))
        indproplist.append(int(ReverseHex(siscontents[14+x+1]),16)+int(ReverseHex(siscontents[11]),16))
        x = x+2
        sisproplist.append(indproplist)
    return sisproplist, SISPropDict

# function SummaryInfoSector() gave us the 'fixed' values,
# but the next sections are of variable lenth. and number.
# The siscontents[13] gives us the number of properties.
# The list sisproplist gives us the property identifier
# and the offset expressed in dec. Dividing this by 4
# gives us the starting chunk
# Output: [['0x1', 184], ['0x2', 192],
# 184 div by 4 = 46  # 192 / 4 = 48
# Therefore property 0x1 (codepage) lies between chunk 46 and 48


def SummaryInfoSector(rawsector):
    # INPUT: List of directory entries. Could be that ONLY ONE has an entry
    SISList = []
    SISList.append("NUMBEROFSECTIONS:")
    SISList.append(len(rawsector))
    for section in rawsector:
        siscontents =  Header2FourBytes(section)
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
        if p == 1:
            SISPropList, SISDocSum = SISOne(siscontents)
            SISList.append(SISDocSum)
            finalsisproplist = sispropitems(SISPropList, siscontents)
            SISList.append(finalsisproplist)
        if p == 2:
            # we likely have a user defined dictionary too..!
            #?? = SISTwo(siscontents)            
            pass
    return SISList, sisproplist, siscontents


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

def dsmain(infile):
    rawstuff = SummaryInfoLocation(infile)
    # note, as there may be two entries rawstuff is now a list of size n
    sisdict, sispropertylist, siscontents = SummaryInfoSector(rawstuff)
    #finalsisproplist = sispropitems(sispropertylist, siscontents)
    return sisdict, finalsisproplist


