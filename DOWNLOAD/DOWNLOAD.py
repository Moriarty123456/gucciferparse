import struct
import binascii
import array
import urllib2
import csv
import os

from hachoir_metadata import metadata
from hachoir_core.cmd_line import unicodeFilename
from hachoir_parser import createParser

####################################
# DOWNLOAD A SINGLE  FILE          #
# SAVES TO downloaded_docs         #
# AND RETURNS LIST OF SOME METADATA#
####################################




def downloadCOMPOUND(url):

    #####################################################
    # USE TO DOWNLOAD A COMPOUND FILE LIKE DOCX OR XLSX #
    # INPUT: url of file                                #
    # OUTPUT: list of filename, type,                   #
    # & server metadata                                 #
    # FILE: saved in /downloaded_docs                   #
    # /doc, or xls/filename                             #
    #####################################################

    infoMeta=[]
    file_name = url.split('/')[-1]
    file_type = file_name.split(".")[-1]
    base_dir = os.path.abspath("../../../downloaded_docs/")
    download_dir = os.path.join(base_dir, file_type)
    infoMeta.append(file_name)
    infoMeta.append(file_type)
    u = urllib2.urlopen(url)
    meta = u.info()
    infoMeta.append(meta.headers)
    doc= u.read()
    f = open(os.path.join(download_dir,file_name), 'wb')
    f.write(doc)
    f.close()    
    print "Done", file_name, " Saved to: ", download_dir
    return infoMeta


def downloadBINARY(url):
    ###########################################################
    # USE TO DOWNLOAD A BINARY FILE LIKE DOC OR XLS           #
    # INPUT: the url of the file.                             #
    # OUTPUT: the hex of the file, and list of some metadata, #
    # from the server and from a hachoir_metadata scan        #
    # SAVES FILE TO: downloaded_docs/doc, or xls/filename     #
    ###########################################################
    infoMeta=[]
    file_name = url.split('/')[-1]
    file_type = file_name.split(".")[-1]
    base_dir = os.path.abspath("../../../downloaded_docs/")
    download_dir = os.path.join(base_dir, file_type)    
    infoMeta.append(file_type)
    infoMeta.append(file_name)
    u = urllib2.urlopen(url)

    meta = u.info()
    infoMeta.append(meta.headers)
    doc= u.read()
    f = open(os.path.join(download_dir,file_name), 'wb')
    f.write(doc)

    with open(os.path.join(download_dir,file_name), 'rb') as p:
    # Slurp the whole file and convert it to hex all at once
        hexdata = binascii.hexlify(p.read())

    # use hachoir to add the standard metadata
    filename = download_dir+ '/'+file_name
    print filename
#    filename = unicodeFilename(filename), filename
    filename, realname = unicodeFilename(filename), filename
    parser = createParser(filename)
    try:
        metalist = metadata.extractMetadata(parser).exportPlaintext()
        infoMeta.append(metalist[1:4])
    except Exception:
        infoMeta.append(["none","none","none"])

        
    p.close()    
    print "Done", file_name, " Saved to: ", download_dir
    return hexdata, infoMeta
