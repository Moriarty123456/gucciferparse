from filetimes import *
import struct
import binascii
import array
import urllib2
import csv

from hachoir_metadata import metadata
from hachoir_core.cmd_line import unicodeFilename
from hachoir_parser import createParser

excel_list = ["1st-quarter-targeted-renewals.xls",
"2016-cycle-passwords.xls","all-money-in-2005-2006.xls","big-donors-list.xls","email_export.xls","hfscmemberdonationsbyparty6101.xls","hsu-contributions.xls","magliochetti-campaign-contributions-99-00.xls","master-spreadsheet-pac-contributions.xls","money-in-2005.xls","out-of-region-ne-donors.xls","paw-family-contributions.xls","pma-group-donations-99-08.xls","pma-group-pac-contributions-99-08.xls","soft-commits.xls","updated-hsu-and-paw-money.xls"]
not_a_list = ["1st-quarter-targeted-renewals.xls"]

def OutputHex(filename):
    with open(filename, "rb") as f:      
        hexdata = binascii.hexlify(f.read())
    return hexdata

def OutputHeader(filename):
    with open(filename, "rb") as f:      
        hexdata = binascii.hexlify(f.read())
        header = hexdata.split("feffffff")[0]
    return header

def Header2FourBytes(header):
    y=8
    headlist=[]
    for x in range(len(header)/8):
        headlist.append(header[(y-8):y])
        y=y+8

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
        if fourbyte[6:8] == "01":
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
    
#####################
# The Magic Happens #
#####################

OffSetChaseList = []
def MainLoop():
    
    #for file in excel_list:
    for file in not_a_list:
        try:
            herx = OutputHex(file)
            header = OutputHeader(file)
            headlist =  Header2FourBytes(header)
            dsllist = ListDSL(headlist)
            root = ListRootOffset(dsllist, herx)
            rootsection, rootin4bytes = RootSectionList(root, herx)
            rtlist=RootTimes(rootin4bytes)
            startingsectorloc, streamsize = StartingSectorLocation(rootin4bytes)
            chaselist = []
            for r in root:

                chaselist.append(file)
                chaselist.append(r)
                chaselist.append(hex(r))
                chaselist.append(headlist)
                #chaselist.append(rootsection)
                chaselist.append(rootin4bytes)
                chaselist.append(rtlist)
                chaselist.append(startingsectorloc)
                chaselist.append(streamsize)
                
            OffSetChaseList.append(chaselist)  
        except:
            print "Computer says 'no'"
    
MainLoop()

for line in OffSetChaseList:
    print "filename: ", line[0]
    print "RootOffset dec / hex: ", line[1], line[2]
    print "Directory Name (64Bytes)", line[4][0:16]
    print "Root Timestamp ", line[5]
    print "Dir Name length, Obj Type, red/black", line[4][16]
    print "Left Sibling ID, (all F's = no sibling)", line[4][17]
    print "Right Sibling ID, (all F's = no sibling)", line[4][18]
    print "Child ID, (all F's = no child)", line[4][19]
    print "CLSID, Root or Storage all zeros", line[4][20:24]
    print "State bits, zeros for root or storage", line[4][24]
    print "Note: The root storage's creation and modification time stamps are normally stored on the file itself in the file system."
    print "Creation time for a storage object. Must be 000's for Root.UTC", line[4][25:27]
    print "Modified time storage object, or root. Stream must be zeros", line[4][27:29]
    print "Starting Sector Location. If stream = sector location. If Root = mini-stream Location. If Storage all zeros", line[4][29]
    print "Starting Sector Size", line[4][30]

    print "starting Sector Location", line[6]
    print "starting Sector Location (hex)", hex(line[6])
    print "sec size & hex", line[7], hex(line[7])
