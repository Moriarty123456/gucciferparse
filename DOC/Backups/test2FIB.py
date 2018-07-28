# coding: utf-8
#!/usr/bin/env python

import sys

sys.path.append("/home/david/Documents/ClintonJail/Guccifer2.0/Docs&Tools/Python/ALL_GUC_FILE_PROJECT/python/modules/DOC")

from binarytoolkit import *
from FIB_BLOB import *

#infile = "/home/david/Documents/ClintonJail/Guccifer2.0/Docs&Tools/Python/ALL_GUC_FILE_PROJECT/downloaded_docs/xls/1st-quarter-targeted-renewals.xls"

infile = "/home/david/Documents/ClintonJail/Guccifer2.0/Docs&Tools/Python/ALL_GUC_FILE_PROJECT/downloaded_docs/doc/old/foreign-donations-fpotus-foundations.doc"


herx = 
base = 
csw =  
fibRgW =  Header2FourBytes(herx[1092:1148]) #(28 bytes)
cslw = herx[1148:1152] #(2 bytes)
fibRgLw =  Header2FourBytes(herx[1152:1328]) #(88 bytes)
cbRgFcLcb =  herx[1328:1332] # (2 bytes)*
startBLOB = 1332
endBLOB = (2*(ReverseInt(cbRgFcLcb)*8))+1332
fibRgFcLcbBlob  =  Header2FourBytes(herx[startBLOB:endBLOB]) # (variable)

#print "is this time?", TimeStamp(fibRgFcLcbBlob[174:176])

fibRgFcLcb97 = fibRgFcLcbBlob[0:186]
fibRgFcLcb2000 = fibRgFcLcbBlob[186:216]
FibRgFcLcb2002 = fibRgFcLcbBlob[216:272]
FibRgFcLcb2003 = fibRgFcLcbBlob[272:328] 

FibRgFcLcb2007 = fibRgFcLcbBlob[328:len(fibRgFcLcbBlob)]

cswNew = herx[endBLOB:endBLOB+4]  #--> 0500
#(2 bytes): An unsigned integer that specifies the count of 16-bit values corresponding to fibRgCswNew that follow. This MUST be one of the following values, depending on the value of nFib.
# Value of nFib cswNew
# 0x00C1 0
# 0x00D9 0x0002
# 0x0101 0x0002
# 0x010C 0x0002
# 0x0112 0x0005

fibRgCswNew = []
CSWSTART = endBLOB+4
count =0
for x in range(ReverseInt(cswNew)):
    fibRgCswNew.append(herx[CSWSTART+count:CSWSTART+4+count])
    count = count +4

nFibNew = fibRgCswNew[0]
# An unsigned integer that specifies the version number of the file format that is used. This value MUST be one of the following.
# Value
# 0x00D9
# 0x0101
# 0x010C
# 0x0112 
if nFibNew == "1201":
    cQuickSavesNew = fibRgCswNew[1]
    lidThemeOther = fibRgCswNew[2]
    lidThemeFE = fibRgCswNew[3]
    lidThemeCS = fibRgCswNew[4]
elif nFibNew == "0c01" or nFibNew == "0101" or nFibNew == "d900":
    cQuickSavesNew = fibRgCswNew[1]
    lidThemeOther = ""
    lidThemeFE = ""
    lidThemeCS = ""
else:
    cQuickSavesNew = ""
    lidThemeOther = ""
    lidThemeFE = ""
    lidThemeCS = ""
    # note: p100 [MS-DOC] --> This value is undefined and MUST be ignored.
#print fibRgCswNew --> ['1201', '0000', '0904', '0000', '0000']
