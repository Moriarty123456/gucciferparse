from filetimes import *
import struct
import binascii
import array
import urllib2

from hachoir_metadata import metadata
from hachoir_core.cmd_line import unicodeFilename
from hachoir_parser import createParser


##################################################################################################################
##### PROJECT TO PARSE DATA INCLUDING BINARY AND SERVER TIMESTAMPS FROM GUCCIFER_2.0. DOCUMENTS ##################
##################################################################################################################

#OPEN CODEPAGES LIST
#CODEPAGE =0042 H = 4200 in hsu.xls at offset 0x299 we see:
#00000280   20 20 20 20  20 20 20 20  20 20 20 20  20 20 20 20
#00000290   20 20 20 20  20 20 20 20  42 00 02 00  B0 04 61 01
#000002A0   02 00 00 00  C0 01 00 00  3D 01 06 00  01 00 02 00

#from pdf excelfileformat.pdf ...
#The CODEPAGE record in BIFF8 always contains the code page 1200
# ... 04B0 H = 1200 = UTF-16 (BIFF8)
#If Size is nonzero and the CodePage property set's CodePage property has the value CP_WINUNICODE (0x04B0), then the value MUST be a null-terminated array of 16-bit Unicode characters, followed by zero padding to a multiple of 4 bytes.


#'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
# CHECK THIS OUT SHARED STRINGS! THIS IS HSU-CONT.XLS
# 000001F0   FF FF FF FF  FF FF FF FF  FF FF FF FF  FF FF FF FF  ................
# 00000200   09 08 10 00  00 06 05 00  D3 18 CD 07  C1 C0 00 00  ................
# 00000210   06 03 00 00  E1 00 02 00  B0 04 C1 00  02 00 00 00  ................
# 00000220   E2 00 00 00  5C 00 70 00  07 00 00 55  53 45 52 2D  ....\.p....USER-
# 00000230   50 43 20 20  20 20 20 20  20 20 20 20  20 20 20 20  PC
# 00000240   20 20 20 20  20 20 20 20  20 20 20 20  20 20 20 20   
# 00000250   20 20 20 20  20 20 20 20  20 20 20 20  20 20 20 20   
# 00000260   20 20 20 20  20 20 20 20  20 20 20 20  20 20 20 20   
# 00000270   20 20 20 20  20 20 20 20  20 20 20 20  20 20 20 20   
# 00000280   20 20 20 20  20 20 20 20  20 20 20 20  20 20 20 20   
# 00000290   20 20 20 20  20 20 20 20  42 00 02 00  B0 04 61 01          B.....a.
# 000002A0   02 00 00 00  C0 01 00 00  3D 01 06 00  01 00 02 00  ........=.......


# THIS IS 1ST QUARTER TARGETED .XLS
# 000001F0   FF FF FF FF  FF FF FF FF  FF FF FF FF  FF FF FF FF  ................
# 00000200   09 08 10 00  00 06 05 00  A0 19 CD 07  C9 C0 00 00  ................
# 00000210   06 03 00 00  E1 00 02 00  B0 04 C1 00  02 00 00 00  ................
# 00000220   E2 00 00 00  5C 00 70 00  07 00 00 55  53 45 52 2D  ....\.p....USER-
# 00000230   50 43 46 6F  72 63 65 20  20 20 20 20  20 20 20 20  PCForce
# 00000240   20 20 20 20  20 20 20 20  20 20 20 20  20 20 20 20
# 00000250   20 20 20 20  20 20 20 20  20 20 20 20  20 20 20 20
# 00000260   20 20 20 20  20 20 20 20  20 20 20 20  20 20 20 20
# 00000270   20 20 20 20  20 20 20 20  20 20 20 20  20 20 20 20
# 00000280   20 20 20 20  20 20 20 20  20 20 20 20  20 20 20 20
# 00000290   20 20 20 20  20 20 20 20  42 00 02 00  B0 04 61 01          B.....a.
# 000002A0   02 00 00 00  C0 01 00 00  3D 01 04 00  01 00 05 00  ........=.......00000230   50
# 000002B0   BA 01 0F 00  0C 00 00 54  68 69 73 57  6F 72 6B 62  .......ThisWorkb


#CODEPAGE IS 42 00 02 00 = 1200 UNICODE
#WHAT IS B0 04 61 01???
#''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''#

#https://msdn.microsoft.com/en-us/library/dd373901(v=vs.85).aspx
#LOCALE_USER_DEFAULT The default locale for the user or process. The value of this constant is 0x0400. 
#LOCALE_SYSTEM_DEFAULT The default locale for the operating system. The value of this constant is 0x0800.
#LOCALE_SABBREVLANGNAME	Abbreviated name of the language. In most cases, the name is created by taking the two-letter language abbreviation from ISO Standard 639 and adding a third letter, as appropriate, to indicate the sublanguage. For example, the abbreviated name for the language corresponding to the English (United States) locale is ENU.
#The value of LOCALE_INVARIANT IS 0x007f.
#Language identifier with a hexadecimal value. For example, English (United States) has the value 0409, which indicates 0x0409 hexadecimal, and is equivalent to 1033 decimal. LISTED IN languagecodes.csv
#LOCALE_IGEOID A 32-bit signed number that uniquely identifies a geographical location. listed in GeoLocCodes.csv

#### FULL_LANG_IDS.CSV HAS THE LOT!

### BIFF versions:
##5.8 BOF – Beginning of File
# BIFF2 0009 H
# BIFF3 0209 H
# BIFF4 0409 H
# BIFF5 0809 H
# BIFF8 0809 H --> has "workbook" stream - Hsu-cont.xls 
# The BOF record is the first record of any kind of stream or substream:


#codepages list is in windows_code_pages.csv. Columns are:
#Identifier,Hex_BE,Hex_BE_py,Hex_LE,Hex_LE_py,DotNET_Name,Name
#eg
#437,01B5,\x01\xb5,B501,\xb5\x01,IBM437,OEM United States
def LE_codepage_query(query):
    # for each codepage in 5th column of windows_code_pages.csv?
    # search for pattern in file
    # typical line 02 00 00 00  E4 04 00 00  1E 00 00 00  08 00 00 00
    # e4 04 = 1252 = english ANSI Latin 1; Western European (Windows)
    # if yes return line(s) where query features, as it may be random in a png for example.
    # and note position (?) as Win32 TS and UUIDs, and important stuff may be close
    pass



# OPEN CSVFILE OF GUCCIFER2.0 ATTACHMENTS, AND DOWNLOAD FILE
outputcsv = "outputfile.csv"
infoMeta=[]

url = "https://guccifer2.files.wordpress.com/2016/07/hsu-contributions.xls"

def Downloadfile(url):

    file_name = url.split('/')[-1]
    infoMeta.append(file_name)
    u = urllib2.urlopen(url)
    f = open(file_name, 'wb')
    meta = u.info()
    infoMeta.append(meta.headers)
    doc= u.read()
    f.write(doc)

    with open(file_name, 'rb') as f:
    # Slurp the whole file and efficiently convert it to hex all at once
        hexdata = binascii.hexlify(f.read())

    # use hachoir to add the standard metadata
    filename = './'+file_name 
    filename, realname = unicodeFilename(filename), filename
    parser = createParser(filename)
    metalist = metadata.extractMetadata(parser).exportPlaintext()

    infoMeta.append(metalist[1:4])

        
    f.close()    
   # print "Done", file_name, " Info is ", infoMeta
    return file_name, hexdata


Downloadfile(url)



#in = raw hex. out = list split at d101 = 2015/16
# note: very basic - have to filter after this
def timestamp(hexdata):
    tstamplistd101=[]
    chunks=hexdata.split("d101")
    for ts in chunks:
        tstamplistd101.append(ts[-12:]+"d101")

    return tstamplistd101


#In = list of 16 char timestamps in BE order. Out= call Win32Time function supplied with list in LE order 
def sortTimeStamps(tslist):
    for ts in tslist:
        infoMeta.append(ts)
        hightime = ts[6:8]+ts[4:6]+ts[2:4]+ts[0:2]
        lowtime =ts[14:16]+ts[12:14]+ts[10:12]+ts[8:10]
        ft= hightime+":"+lowtime
        infoMeta.append(ft)
        win32time(ft)

sortTimeStamps(timestamp(hexdata))

def win32time(ft):
    # switch parts
    h2, h1 = [int(h, base=16) for h in ft.split(':')]
    # rebuild
    ft_dec = struct.unpack('>Q', struct.pack('>LL', h1, h2))[0]
    # add ft_dec to infometa
    infoMeta.append(ft_dec)
    # use function from iceaway's comment
    infoMeta.append(filetime_to_dt(ft_dec))


print infoMeta

    
#win32time(ft)
