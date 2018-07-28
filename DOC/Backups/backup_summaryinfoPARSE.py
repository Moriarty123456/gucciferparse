import sys

sys.path.append("/home/david/Documents/ClintonJail/Guccifer2.0/Docs&Tools/Python/ALL_GUC_FILE_PROJECT/python/modules/DOC")

from binarytoolkit import *

def SummaryInfoLocation(infile):
    herx = OutputHex(infile)    
    sis ="05005300"+ herx.split("05005300")[1][0:248]
    sislist =  Header2FourBytes(sis)
    sl= sislist[-3]
    secloc = sl[6:8]+sl[4:6]+sl[2:4]+sl[0:2]
    # x 2 for bits
    decimalsectorlocation = 2*((int(secloc,16)+1)*512)
    endsidlocation = decimalsectorlocation + 1024
    print decimalsectorlocation, endsidlocation
    

    return herx[decimalsectorlocation:endsidlocation]

def ReverseHex(somehex):
    return hex(int(somehex[6:8]+somehex[4:6]+somehex[2:4]+somehex[0:2],16))

def SummaryInfoSector(rawsector):
    #see https://msdn.microsoft.com/en-us/library/dd944893(v=office.12).aspx
    SISDict = {}
    siscontents =  Header2FourBytes(rawsector)
    SISDict["WORDbyteOrder"]= siscontents[0][:4]
    SISDict["WORD-version"] = siscontents[0][4:]
    SISDict["OSMajorVersion"] = siscontents[1][:2]
    SISDict["OSMinorVersion"] = siscontents[1][2:4]
    SISDict["OSType"] = siscontents[1][4:8]
    SISDict["GUID-applicationClsid"] = siscontents[2]+siscontents[3]+siscontents[4]+siscontents[5]
    SISDict["DWORD-cSections"] = siscontents[6]
    SISDict["GUID-formatId"] = siscontents[7]+siscontents[8]+siscontents[9]+siscontents[10]
    sl = siscontents[11]
    SISDict["FilePointer-sectionOffset"] = hex(int(sl[6:8]+sl[4:6]+sl[2:4]+sl[0:2],16))
    sz = siscontents[12]
    SISDict["FilePointer-SectionSize"] = hex(int(sz[6:8]+sz[4:6]+sz[2:4]+sz[0:2],16))
    num = siscontents[13]
    SISDict["NumberPropertiesHex"] = hex(int(num[6:8]+num[4:6]+num[2:4]+num[0:2],16))
    SISDict["NumberPropertiesDec"] = int(num[6:8]+num[4:6]+num[2:4]+num[0:2],16)
    npd = int(num[6:8]+num[4:6]+num[2:4]+num[0:2],16)
    sisproplist = []
    x = 0
    while x < npd*2:
        indproplist = []
        indproplist.append(ReverseHex(siscontents[14+x]))
        indproplist.append(int(ReverseHex(siscontents[14+x+1]),16)+int(ReverseHex(siscontents[11]),16))
        x = x+2
        sisproplist.append(indproplist)
    return SISDict, sisproplist, siscontents

# function SummaryInfoSector() gave us the 'fixed' values,
# but the next sections are of variable lenth. and number.
# The siscontents[13] gives us the number of properties.
# The list sisproplist gives us the property identifier
# and the offset expressed in dec. Dividing this by 4
# gives us the starting chunk
# Output: [['0x1', 184], ['0x2', 192],
# 184 div by 4 = 46  # 192 / 4 = 48
# Therefore property 0x1 (codepage) lies between chunk 46 and 48


def sispropitems(sisproplist, siscontents):
    ############################################################
    # INPUT: the list of properties ("0x1", 184),              #
    # INPUT: the contents of the Summary info stream in chunks #
    # OUTPUT: Dictionary of meaning and value.                 #
    ############################################################
    theproperties = [("GKPIDSI_CODEPAGE", "01000000", "0x1", "02000000", "no", "4"), ("GKPIDSI_TITLE", "02000000", "0x2", "1E000000", "yes", "ANSI"), ("GKPIDSI_SUBJECT", "03000000", "0x3", "1E000000", "yes", "ANSI"), ("GKPIDSI_AUTHOR", "04000000", "0x4", "1E000000", "yes", "ANSI"), ("GKPIDSI_KEYWORDS", "05000000", "0x5", "1E000000", "yes", "ANSI"), ("GKPIDSI_COMMENTS", "06000000", "0x6", "1E000000", "yes", "ANSI"), ("GKPIDSI_TEMPLATE", "07000000", "0x7", "1E000000", "yes", "ANSI"), ("GKPIDSI_LASTAUTHOR", "08000000", "0x8", "1E000000", "yes", "ANSI"), ("GKPIDSI_REVNUMBER", "09000000", "0x9", "1E000000", "yes", "4"), ("GKPIDSI_EDITTIME", "0A000000", "0xa", "40000000", "no", "8"), ("GKPIDSI_LASTPRINTED", "0B000000", "0xb", "40000000", "no", "8"), ("GKPIDSI_CREATE_DTM", "0C000000", "0xc", "40000000", "no", "8"), ("GKPIDSI_LASTSAVE_DTM", "0D000000", "0xd", "40000000", "no", "8"), ("GKPIDSI_PAGECOUNT", "0E000000", "0xe", "03000000", "no", "4"), ("GKPIDSI_WORDCOUNT", "0F000000", "0xf", "03000000", "no", "4"), ("GKPIDSI_CHARCOUNT", "10000000", "0x10", "03000000", "no", "4"), ("GKPIDSI_THUMBNAIL", "11000000", "0x11", "47000000", "no", "Data"), ("GKPIDSI_APPNAME", "12000000", "0x12", "1E000000", "yes", "ANSI"), ("GKPIDSI_DOC_SECURITY", "13000000", "0x13", "03000000", "no", "4")]
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

infile = "/home/david/Documents/ClintonJail/Guccifer2.0/Docs&Tools/Python/ALL_GUC_FILE_PROJECT/downloaded_docs/xls/out-of-region-ne-donors.xls"

# 53760, 4096]
# 57856, 4096]
def simain(infile):
    rawstuff = SummaryInfoLocation(infile)
    #print rawstuff
    sisdict, sispropertylist, siscontents = SummaryInfoSector(rawstuff)
    print sisdict
    finalsisproplist = sispropitems(sispropertylist, siscontents)
    return sisdict, finalsisproplist

simain(infile)
