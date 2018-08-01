import sys
import struct
import binascii
import array
import urllib2
import csv
from itertools import chain


from hachoir_metadata import metadata
from hachoir_core.cmd_line import unicodeFilename
from hachoir_parser import createParser

sys.path.append("/home/david/Documents/ClintonJail/Guccifer2.0/Docs&Tools/Python/ALL_GUC_FILE_PROJECT/python/modules/DOC")

from filetimes import *


def OutputHex(filename):

    with open(filename, "rb") as f:      
        hexdata = binascii.hexlify(f.read())
    f.close()

    return hexdata

def OutputHeader(herx):
    header = herx.split("feffffff")[0]
    return header


def OutputFFslist(herx, string, stringlength):
    # INPUT: result of OutputHex (long string of hex) and string to find in it, i.e OutputFFslist(herx,"05004400", 256) & length to split at
    # OUTPUT: two lists of position and chunks
    positions = []
    ChunkList = []
    # to account for the length of the string that is searched for delete that length from string, so we can just search a location with stringlength = 256

    thelength = stringlength-len(string)
    Listwithoutstring = herx.split(string)[1:]
    for x in Listwithoutstring:
        ChunkList.append(string+x[0:thelength])
    y = find_all(herx, string)
    for t in y:
        positions.append(t/2)
    return positions, ChunkList

def PositionChunkList(herx, string, stringlength):
    # Mostly copy of OutputFFslist, but for more general use with one list as output
    # INPUT: result of OutputHex (long string of hex) and string to find in it, i.e OutputFFslist(herx,"05004400", 256) & length to split at
    # OUTPUT: onelist position and chunks [position1,chunk1,positoin2,chunk2...]
    positions = []
    ChunkList = []
    # to account for the length of the string that is searched for delete that length from string, so we can just search a location with stringlength = 256

    thelength = stringlength-len(string)
    Listwithoutstring = herx.split(string)[1:]
    for x in Listwithoutstring:
        ChunkList.append(string+x[0:thelength])
    y = find_all(herx, string)
    for t in y:
        positions.append(t/2)
    poschunklist = list(chain(*zip(positions,ChunkList)))
    return poschunklist



def find_all(a_str, sub):
    start = 0
    while True:
        start = a_str.find(sub, start)
        if start == -1: return
        yield start
        start += len(sub) # use start += 1 to find overlapping matches
        
def Header2FourBytes(header):
    y=8
    headlist=[]
    for x in range(len(header)/8):
        headlist.append(header[(y-8):y])
        y=y+8

    return headlist

def HexToEightBytes(herx):
    y=16
    headlist=[]
    for x in range(len(herx)/16):
        headlist.append(herx[(y-16):y])
        y=y+16

    return headlist



# Directory Sector Locations (DSL's) start at the Ninth 4 byte chunk
def ListDSL(headlist):
    DSLList=[]
    for ds in headlist[9:]:
         FourByte = ds[6:8]+ds[4:6]+ds[2:4]+ds[0:2]
         if int(FourByte,16) == 0:
             pass
         else:
             SL = (int(FourByte,16)+1)*512
             DSLList.append(SL)

    return DSLList
        
def ListRootOffset(headlist, herx):
    RootOffsetList =[]
    for offset in headlist:
        #note herx offsets will be in bits, so x2 for Bytes
        if "52006f006f00740020" in herx[(2*offset):((2*offset)+48)]:
            RootOffsetList.append(offset)
        else:
            pass
    return RootOffsetList

def RootSectionList(rootoffsetlist, herx):
    for root in rootoffsetlist:
        section = herx[(2*root):(2*(root+256))]
        root4bytelist = Header2FourBytes(section)
    return section, root4bytelist

def StartingSectorLocation(root4bytelist):
    # a compound File Directory, like the Root Directory is 128 Bytes long. The last two 4 byte chunks (29 & 30) will be the Starting Sector Location (if it exists), and the Stream Size. As before we add one to the number then times by the sector size.
    sl = root4bytelist[29]
    secloc = sl[6:8]+sl[4:6]+sl[2:4]+sl[0:2]
    sectorlocation = (int(secloc,16)+1)*512
    ss = root4bytelist[30]
    streamsize = int((ss[6:8]+ss[4:6]+ss[2:4]+ss[0:2]),16)
    return sectorlocation, streamsize
    
def RootTimes(root4bytelist):

    rtimelist=[]
    index = 0
    indlist=[]
    ftimelist = []
    # if the 4 byte section ends in 01 it could be a timestamp, but not if the previous 4 Byte section is all zeros.
    for fourbyte in root4bytelist:
        if fourbyte[6:8] == "01" or fourbyte[6:8] == "02":
            indlist.append(index)
        index = index + 1
    
    # Check to see if prior chunk is just zeros
    for inchunk in indlist:
        if int(root4bytelist[inchunk-1],16) == 0:
            pass
        else:
            timestamp = root4bytelist[inchunk-1] + root4bytelist[inchunk]
            rtimelist.append(timestamp)
    for ts in rtimelist:
        hightime = ts[6:8]+ts[4:6]+ts[2:4]+ts[0:2]
        lowtime =ts[14:16]+ts[12:14]+ts[10:12]+ts[8:10]
        ft= hightime+":"+lowtime
        h2, h1 = [int(h, base=16) for h in ft.split(':')]
        ft_dec = struct.unpack('>Q', struct.pack('>LL', h1, h2))[0]
        naive = (filetime_to_dt(ft_dec))
        filetime = naive.isoformat()
        ftimelist.append(filetime)
        
    return ftimelist

def TimeStamp(twobyfourbytelist):
    ts = "".join(twobyfourbytelist)
    hightime = ts[6:8]+ts[4:6]+ts[2:4]+ts[0:2]
    lowtime =ts[14:16]+ts[12:14]+ts[10:12]+ts[8:10]
    ft= hightime+":"+lowtime
    h2, h1 = [int(h, base=16) for h in ft.split(':')]
    ft_dec = struct.unpack('>Q', struct.pack('>LL', h1, h2))[0]
    naive = (filetime_to_dt(ft_dec))
    filetime = naive.isoformat()

    return filetime

def EditTime(twobyfourbytelist):
    ts ="".join(twobyfourbytelist)
    hightime = ts[6:8]+ts[4:6]+ts[2:4]+ts[0:2]
    lowtime =ts[14:16]+ts[12:14]+ts[10:12]+ts[8:10]
    ft= hightime+":"+lowtime
    h2, h1 = [int(h, base=16) for h in ft.split(':')]
    ft_dec = struct.unpack('>Q', struct.pack('>LL', h1, h2))[0]
    editsecs = (ft_dec/1000000000)
        
    return editsecs


def SummaryInfoLocation(file):
    herx = OutputHex(file)    
    sis ="05005300"+ herx.split("05005300")[1][0:248]
    sislist =  Header2FourBytes(sis)
    sl= sislist[-3]
    secloc = sl[6:8]+sl[4:6]+sl[2:4]+sl[0:2]
    # x 2 for bits
    decimalsectorlocation = 2*((int(secloc,16)+1)*512)
    endsidlocation = decimalsectorlocation + 1024

    return herx[decimalsectorlocation:endsidlocation]

def ReverseHex(somehex):
    return hex(int(somehex[6:8]+somehex[4:6]+somehex[2:4]+somehex[0:2],16))

def ReverseInt(somehex):
    if len(somehex) == 8:
        return int(somehex[6:8]+somehex[4:6]+somehex[2:4]+somehex[0:2],16)
    elif len(somehex) == 4:
        return int(somehex[2:4]+somehex[0:2],16)
    elif len(somehex) == 2:
        return int(somehex,16)
    else:
        giveago = "".join(somehex)
        return int(giveago,16)

def ReverseText(somehex):
    return somehex[6:8]+somehex[4:6]+somehex[2:4]+somehex[0:2]

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

def main(infile):
    rawstuff = SummaryInfoLocation(infile)
    sisdict, sispropertylist, siscontents = SummaryInfoSector(rawstuff)
    finalsisproplist = sispropitems(sispropertylist, siscontents)


def checkLocale(fourbytehexlocale):
    #input is say string "09040000" little endian, 0409 big endian = 1033 = US English
    locales = [("09040000", "US English"), ("19040000", "Russia Russian"), ("19000000", "Russian other"), ("22040000", "Ukraine Ukrainian"), ("00010000", "Other"), ("22000000", "Ukrainian Lang"), ("09000000", "English Lang"), ("18000000", "Romanian Lang"), ("18080000", "Romainian Moldova"), ("18040000", "Romanian Romania")]
    ans = str(fourbytehexlocale)
    for entry in locales:
        if entry[0] == fourbytehexlocale:
            ans = entry[1]
        else:
            ans = ans
    return ans


def olddicktophone(dick):
    #input a dictionary - the bit given by Dictionary identifier 00000000
    # eg dick = ['05000000', '02000000', '14000000', '5f004100', '64004800', '6f006300', '52006500', '76006900', '65007700', '43007900', '63006c00', '65004900', '44000000 ...
    #output the text of those fields with number
    # eg {2: '_AdHocReviewCycleID', 3: '_NewReviewCycle', 4: '_EmailSubject', 5: '_AuthorEmail', 6: '_AuthorEmailDisplayName'}
    numof = int(ReverseHex(dick[0]),16)
    mydick = {}
    #posn is counter of position
    posn = 2
    startnum = int(ReverseHex(dick[posn-1]),16)
    #startlength = (int(ReverseHex(dick[posn+1]),16))/2
    #print startlength
    itemno = 0
    for x in range(numof):
        itemno = x + startnum
        #  print "item no is ", itemno,
        # hack to solve problem about modulus
        t = (int(ReverseHex(dick[posn]),16))
        if t/2 == (t-1)/2:
            length = (t+1)/2
        else:
            length = t/2
            #    print "length for ", dick[posn], "is ", length
            #   print "result ", dick[posn+1:posn+1+length]

        thetext = "".join(dick[posn+1:posn+1+length]).decode("hex")
        if "\x00" in thetext:
            thetext = "".join(thetext.split("\x00"))
        else:
            pass
        mydick[itemno] = thetext
        posn = posn + (length+2)
    return mydick

def hextoascii(chunks):
    #chunks as list or string
    if type(chunks) == list:
        thetext = "".join(chunks).decode("hex")
    else:
        thetext = chunks.decode("hex")
    if "\x00" in thetext:
        thetext = "".join(thetext.split("\x00"))
    else:
        thetext = thetext

    return thetext


def numGuids(siscontents):
     # some document summaries have two guids. one for document summary and one for user defined info. The number of guids is given by DWORD-cSections.
    n = 2
    x = siscontents[6]
    p=int([x[i:i+n] for i in range(0, len(x), n)][0])
    # p returns the correct number of guids.
    return p


def dicktophone(dick):
    # The problem is that sometimes the document splits a chunk. i.e normally we would have (in chunks):
    # 08 00 00 00  02 00 00 00  0C 00 00 00  5F 50 49 44 ...
    # [num props] [dict entry num] [length] [dict entry]...
    # but we also get:
    # 74 00 *06* 00  00 00 0D 00  00 00 5F 41  75 74 68 6F...
    # [num 06 starts in middle of chunk] [length in middle of chunk] [dict starts in middle of chunk] ...
    # so if dicktophone fails we have to join the chunks and work in bytes
    # we *assume* the first two chunks are normal...
    mydick = {}
    dickbytes = "".join(dick)
    numdict = int(ReverseHex(dickbytes[0:8]),16)
    startnum = int(ReverseHex(dickbytes[8:16]),16)
    #then we shift to bytes...
    positionsindickt = []
    for x in range(numdict):
        itemnum = "0"+str(x+startnum)+"000000"
        itemlength = dickbytes[8:].split(itemnum)[1][0:8]
        declength = int(ReverseHex(itemlength),16)
        itemconts = dickbytes[8:].split(itemnum)[1][8:8+(2*declength)]
        mydick[x+startnum] = hextoascii(itemconts)        
    return mydick

def PIDHlinks(inhex):
    #input the contents in chunks of _PID_HLINKS
    #Output list of emails
    PIDEmails = []
    wtype = inhex[0]
    cbData = ReverseInt(inhex[1])
    cElements = ReverseInt(inhex[2])
    numberElements = cElements/6
    linkssection = "".join(inhex[3:])
    rawemaillist = linkssection.split("1f0000000100000000000000")[:-1]
    for x in rawemaillist:
        anemail =  hextoascii(x[80:])
        PIDEmails.append(anemail)
    if len(rawemaillist) == numberElements:
        return PIDEmails
    else:
        print "PROBLEM with PIDHlinks"
        return hextoascii(rawemaillist)

def hextoBINARYlist(somehex, num_of_bits):
    binlist = []
    scale = 16 ## equals to hexadecimal
    thebinary = bin(int(somehex, scale))[2:].zfill(num_of_bits)
    for b in thebinary:
        binlist.append(b)
    return binlist

def DTTMfromHex(th):
    # page 238 MSDOC
    # h is 4bytes of hex.
    #th = ReverseText(h)
    ts = bin(int(th,16))[2:].zfill(32)
    # ex: '100001-10110-01100-1001-001010011-011'
    mins = int(ts[:6],2) #6 bits
    hr = int(ts[6:11],2) #5
    dom = int(ts[11:16],2) #5
    mon = int(ts[16:20],2) #4
    yr = int(ts[20:29],2)+1900 #9 +1900
    wdy = int(ts[29:32],2)#3 sun == 0
    #2015-02-18T17:16:01.422000
    return str(yr) + "-" + str(mon).zfill(2) + "-" + str(dom).zfill(2) + "T" + str(hr) + ":" + str(mins) + ":00"


def Downloadfile(url):
    #INPUT: the url of the file
    #OUTPUT:  file written to disk and the server metadata
    infoMeta = []
    file_name = url.split('/')[-1]
    infoMeta.append(file_name)
    u = urllib2.urlopen(url)

    meta = u.info()
    infoMeta.append(meta.headers)
    doc= u.read()
    f = open(file_name, 'wb')
    f.write(doc)
    # use hachoir to add the standard metadata
    filename = './'+file_name
    filename, realname = unicodeFilename(filename), filename
    parser = createParser(filename)
    try:
        metalist = metadata.extractMetadata(parser).exportPlaintext()
        infoMeta.append(metalist[1:4])
    except Exception:
        infoMeta.append(["none","none","none"])
    p.close()    
    return file_name, infoMeta

def writelinetoCSV(aline):
    #INPUT: a list with items to write to CSV
    #OUTPUT: none. Line written to CSV
    with open("mydocs.csv", "w") as f:
        an_entry = csv.writer(f)
        an_entry.writerow(aline)
    f.close()

def dummy(teststuff):
    print "version control"

    
