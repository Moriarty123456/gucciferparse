import sys

sys.path.append("/home/david/Documents/ClintonJail/Guccifer2.0/Docs&Tools/Python/ALL_GUC_FILE_PROJECT/python/modules/DOC")

from binarytoolkit import *
from docsuminfoPARSE import *

infile = "/home/david/Documents/ClintonJail/Guccifer2.0/Docs&Tools/Python/ALL_GUC_FILE_PROJECT/downloaded_docs/xls/1st-quarter-targeted-renewals.xls"

def SummaryInfoLocation(infile):
    herx = OutputHex(infile)
    # some documents have two directory entries, so make list of directory entries.
    herxlist = herx.split("05004400")
    docInfoSectorList = []

    print herxlist[1][0:10]
    print herxlist[2][0:10]
    for x in herxlist:
        if x[0:10] == "6f00630075":
            sis ="05004400"+ x[0:248]
            #4ca00
            sislist =  Header2FourBytes(sis)
            # Location of Doc Sum Info is given by third to last chunk
            sl= sislist[-3]
            # check that this location is not blank:
            if int(sl,16) == 0:
                pass
            else:
                secloc = sl[6:8]+sl[4:6]+sl[2:4]+sl[0:2]
                # Size of Doc Sum Info is given by second to last chunk. Seems to usually be 00100000, or 4096 in dec.
                ss = sislist[-2]
                secsize = ss[6:8]+ss[4:6]+ss[2:4]+ss[0:2]
                # Calculate start and end of sector & append to list.
                decimalsectorlocation = 2*((int(secloc,16)+1)*512)
                endsidlocation = decimalsectorlocation + 2*(int(secsize,16))
                docInfoSectorList.append(herx[decimalsectorlocation:endsidlocation])
                #returns number of Summary Info Sectors, and the sector itself
    return docInfoSectorList


SummaryInfoLocation(infile)
