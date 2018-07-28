from filetimes import *
import struct
import binascii
import array
import urllib2
import csv

from hachoir_metadata import metadata
from hachoir_core.cmd_line import unicodeFilename
from hachoir_parser import createParser


###########################################################################
###########                                              ##################
########### PARSE ALL XLS FILES. EXTRACT ALL REGULAR     ##################
########### METADATA TIMESTAMPS, PLUS SERVER TIMESTAMPS, ##################
########### PLUS ANY 2016 TIMESTAMPS FROM GUCCIFER_2.0.  ##################
########### DOCUMENTS                                    ##################
###########                                              ##################
###########################################################################
# OPEN CSVFILE OF GUCCIFER2.0 ATTACHMENTS, AND DOWNLOAD FILE
outputcsv = "outputfile.csv"
infoMeta=[]


def Downloadfile(url):

    file_name = url.split('/')[-1]
    infoMeta.append(file_name)
    u = urllib2.urlopen(url)

    meta = u.info()
    print "Meta is: ", meta
    infoMeta.append(meta.headers)
    doc= u.read()
    f = open(file_name, 'wb')
    f.write(doc)

    with open(file_name, 'rb') as f:
    # Slurp the whole file and efficiently convert it to hex all at once
        hexdata = binascii.hexlify(f.read())

    # use hachoir to add the standard metadata
    filename = './'+file_name 
    filename, realname = unicodeFilename(filename), filename
    parser = createParser(filename)
    try:
        metalist = metadata.extractMetadata(parser).exportPlaintext()
        infoMeta.append(metalist[1:4])
    except Exception:
        infoMeta.append(["none","none","none"])

        
    f.close()    
   # print "Done", file_name, " Info is ", infoMeta
    return file_name, hexdata


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
        return ft



def win32time(ft):
    # switch parts
    h2, h1 = [int(h, base=16) for h in ft.split(':')]
    # rebuild
    ft_dec = struct.unpack('>Q', struct.pack('>LL', h1, h2))[0]
    # add ft_dec to infometa
    infoMeta.append(ft_dec)
    # use function from iceaway's comment
    infoMeta.append(filetime_to_dt(ft_dec))


# to do get the main stuff a workin'



def Main(url):
    print "Really about to dl: ", url
    file_name, hexdata = Downloadfile(url)
    tstamplistd101 = timestamp(hexdata)
    ft = sortTimeStamps(tstamplistd101)
    win32time(ft)
    with open("mycsv.csv", "w") as f:
        an_entry = csv.writer(f)
        an_entry.writerow([infoMeta[0], infoMeta[1][5], infoMeta[1][3], infoMeta[2][0], infoMeta[2][1], infoMeta[2][2], infoMeta[3:]])
    f.close()


#url = "https://guccifer2.files.wordpress.com/2016/07/hsu-contributions.xls"
#Main(url)

with open("attachments.csv", "r") as atts:
    attline = csv.reader(atts)

    for line in attline:
        if "xls" in line[1] and "xlsx" not in line[1]:
            print "Downloading: ", line[1]
            Main(line[1])
        else:
            pass
    atts.close()


    

