import zipfile
import os
import csv
import urllib2

# INPUT: path to the zip, docx, or xlsx file, 
# OUTPUT: New folder based on filename. Files should be in it
def unzip_it(path):
    directory = "/tmp/DOCPROJECT/"+(path.split(".")[0]).split("/")[-1]
    # if not os.path.exists(directory):
    #      os.makedirs(directory)
    zip_ref = zipfile.ZipFile(path, 'r')
    zip_ref.extractall(directory)
    zip_ref.close()
    return directory



# INPUT: The directory
# OUTPUT: list of files, with full paths
def directorywalk(extracted_directory):
    filelist =[]
    for root, dirs, files in os.walk(extracted_directory, topdown=True):
        for name in files:
            s = os.path.join(root, name)
            filelist.append(s)
            #print os.path.getsize(s), s
    return filelist


# INPUT: the name of the file like /docProps/app.xml
# OUTPUT: list split at "><" with xml fields.      
def readNprint(props):
    thexml = open(props, 'r')
    textxml = thexml.read()
    thexml.close()
    xmlsplit = textxml.split("><")
    return xmlsplit
    # for x in xmlsplit:
    #     if "http" in x:
    #         pass
    #     else:
    #         print x




def ListOfStrings():
    stringlist = []
    stuff = open("strings.txt", "r")
    slist = (stuff.read()).split(",")
    stuff.close()
    for s in slist:
        stringlist.append(s)
    return stringlist
#############
# MAIN LOOP #
#############

def DocLoop(path_to_zip_file):
    print "Looping ", path_to_zip_file

    file_directory = unzip_it(path_to_zip_file)
    doc_file_names = directorywalk(file_directory)
#    no_files = len(doc_file_names)
    strlist = ListOfStrings()
    filecount = 0
    
    for name in doc_file_names:
        littlelist = []
        bits = readNprint(name)
        fName = name.split("/tmp/DOCPROJECT/")[1]
        littlelist.append(fName)
        for item in strlist:
            for b in bits:
                if item in b:
                    littlelist.append(b)
                else:
                    pass
        with open('test.csv', 'ab') as fp:
            wr = csv.writer(fp, dialect='excel')
            wr.writerow(littlelist)


#backup
    # for item in strlist:
    #     for name in doc_file_names:
    #         bits = readNprint(name)
    #         for b in bits:
    #             littlelist = []
    #             if item in b:
    #                 with open('test.csv', 'ab') as fp:
    #                     wr = csv.writer(fp, dialect='excel')
    #                     wr.writerow([name,item,b])
    #             else:
    #                 pass
                

def Downloadfile(url):

    directory = "/tmp/DOCPROJECT/"+(url.split("/")[-1]).split(".")[0]
    if not os.path.exists(directory):
         os.makedirs(directory)
    file_name = url.split('/')[-1]
    new_file = os.path.join(directory, file_name)
    u = urllib2.urlopen(url)
    doc= u.read()
    f = open(new_file, 'wb')
    f.write(doc)
        
    f.close()    
    return new_file

#attachmenst in /tmp/DOCPROJECT/
with open("attachments.csv", "r") as atts:
    attline = csv.reader(atts)

    for line in attline:
        if "docx" in line[1] or "xlsx" in line[1]:
            print "Downloading file ", line[1]
            DocLoop(Downloadfile(line[1]))
            
        else:
            pass
    atts.close()







# from filetimes import *
# import struct
# import binascii
# import array
# import urllib2
# import csv




# from hachoir_metadata import metadata
# from hachoir_core.cmd_line import unicodeFilename
# from hachoir_parser import createParser

# excel_list = ["1st-quarter-targeted-renewals.xls",
# "2016-cycle-passwords.xls","all-money-in-2005-2006.xls","big-donors-list.xls","email_export.xls","hfscmemberdonationsbyparty6101.xls","hsu-contributions.xls","magliochetti-campaign-contributions-99-00.xls","master-spreadsheet-pac-contributions.xls","money-in-2005.xls","out-of-region-ne-donors.xls","paw-family-contributions.xls","pma-group-donations-99-08.xls","pma-group-pac-contributions-99-08.xls","soft-commits.xls","updated-hsu-and-paw-money.xls"]
# not_a_list = ["1st-quarter-targeted-renewals.xls"]

# def OutputHex(filename):
#     with open(filename, "rb") as f:      
#         hexdata = binascii.hexlify(f.read())
#     return hexdata

# def OutputHeader(filename):
#     with open(filename, "rb") as f:      
#         hexdata = binascii.hexlify(f.read())
#         header = hexdata.split("feffffff")[0]
#     return header

# def Header2FourBytes(header):
#     y=8
#     headlist=[]
#     for x in range(len(header)/8):
#         headlist.append(header[(y-8):y])
#         y=y+8

#     return headlist

# # Directory Sector Locations (DSL's) start at the Ninth 4 byte chunk
# def ListDSL(headlist):
#     DSLList=[]
#     for ds in headlist[9:]:
#          FourByte = ds[6:8]+ds[4:6]+ds[2:4]+ds[0:2]
#          if int(FourByte,16) == 0:
#              pass
#          else:
#              SL = (int(FourByte,16)+1)*512
#              DSLList.append(SL)

#     return DSLList
        
# def ListRootOffset(headlist, herx):
#     RootOffsetList =[]
#     for offset in headlist:
#         #note herx offsets will be in bits, so x2 for Bytes
#         if "52006f006f00740020" in herx[(2*offset):((2*offset)+48)]:
#             RootOffsetList.append(offset)
#         else:
#             pass
#     return RootOffsetList

# def RootSectionList(rootoffsetlist, herx):
#     for root in rootoffsetlist:
#         section = herx[(2*root):(2*(root+256))]
#         root4bytelist = Header2FourBytes(section)
#     return section, root4bytelist

# def StartingSectorLocation(root4bytelist):
#     # a compound File Directory, like the Root Directory is 128 Bytes long. The last two 4 byte chunks (29 & 30) will be the Starting Sector Location (if it exists), and the Stream Size. As before we add one to the number then times by the sector size.
#     sl = root4bytelist[29]
#     secloc = sl[6:8]+sl[4:6]+sl[2:4]+sl[0:2]
#     sectorlocation = (int(secloc,16)+1)*512
#     ss = root4bytelist[30]
#     streamsize = int((ss[6:8]+ss[4:6]+ss[2:4]+ss[0:2]),16)
#     return sectorlocation, streamsize
    
# def RootTimes(root4bytelist):

#     rtimelist=[]
#     index = 0
#     indlist=[]
#     ftimelist = []
#     # if the 4 byte section ends in 01 it could be a timestamp, but not if the previous 4 Byte section is all zeros.
#     for fourbyte in root4bytelist:
#         if fourbyte[6:8] == "01":
#             indlist.append(index)
#         index = index + 1
    
#     # Check to see if prior chunk is just zeros
#     for inchunk in indlist:
#         if int(root4bytelist[inchunk-1],16) == 0:
#             pass
#         else:
#             timestamp = root4bytelist[inchunk-1] + root4bytelist[inchunk]
#             rtimelist.append(timestamp)
#     for ts in rtimelist:
#         hightime = ts[6:8]+ts[4:6]+ts[2:4]+ts[0:2]
#         lowtime =ts[14:16]+ts[12:14]+ts[10:12]+ts[8:10]
#         ft= hightime+":"+lowtime
#         h2, h1 = [int(h, base=16) for h in ft.split(':')]
#         ft_dec = struct.unpack('>Q', struct.pack('>LL', h1, h2))[0]
#         naive = (filetime_to_dt(ft_dec))
#         filetime = naive.isoformat()
#         ftimelist.append(filetime)
        
#     return ftimelist
    
# #####################
# # The Magic Happens #
# #####################

# OffSetChaseList = []
# def MainLoop():
    
#     #for file in excel_list:
#     for file in not_a_list:
#         try:
#             herx = OutputHex(file)
#             header = OutputHeader(file)
#             headlist =  Header2FourBytes(header)
#             dsllist = ListDSL(headlist)
#             root = ListRootOffset(dsllist, herx)
#             rootsection, rootin4bytes = RootSectionList(root, herx)
#             rtlist=RootTimes(rootin4bytes)
#             startingsectorloc, streamsize = StartingSectorLocation(rootin4bytes)
#             chaselist = []
#             for r in root:

#                 chaselist.append(file)
#                 chaselist.append(r)
#                 chaselist.append(hex(r))
#                 chaselist.append(headlist)
#                 #chaselist.append(rootsection)
#                 chaselist.append(rootin4bytes)
#                 chaselist.append(rtlist)
#                 chaselist.append(startingsectorloc)
#                 chaselist.append(streamsize)
                
#             OffSetChaseList.append(chaselist)  
#         except:
#             print "Computer says 'no'"
    
# MainLoop()

# for line in OffSetChaseList:
#     print "filename: ", line[0]
#     print "RootOffset dec / hex: ", line[1], line[2]
#     print "Directory Name (64Bytes)", line[4][0:16]
#     print "Root Timestamp ", line[5]
#     print "Dir Name length, Obj Type, red/black", line[4][16]
#     print "Left Sibling ID, (all F's = no sibling)", line[4][17]
#     print "Right Sibling ID, (all F's = no sibling)", line[4][18]
#     print "Child ID, (all F's = no child)", line[4][19]
#     print "CLSID, Root or Storage all zeros", line[4][20:24]
#     print "State bits, zeros for root or storage", line[4][24]
#     print "Note: The root storage's creation and modification time stamps are normally stored on the file itself in the file system."
#     print "Creation time for a storage object. Must be 000's for Root.UTC", line[4][25:27]
#     print "Modified time storage object, or root. Stream must be zeros", line[4][27:29]
#     print "Starting Sector Location. If stream = sector location. If Root = mini-stream Location. If Storage all zeros", line[4][29]
#     print "Starting Sector Size", line[4][30]

#     print "starting Sector Location", line[6]
#     print "starting Sector Location (hex)", hex(line[6])
#     print "sec size & hex", line[7], hex(line[7])


# ################################
# # GOTCHA YOU LITTLE FUCKERS!   #
# # 1ST-TARGETED-RENEWALS SECTOR #
# # LOCATION 0x4FE00 SHOWS LANG  #
# # SETTINGS E404 = US ENGLISH   #
# ################################





# # #in = raw hex. out = list split at d101 = 2015/16
# # # note: very basic - have to filter after this
# # def timestamp(hexdata):
# #     tstamplistd101=[]
# #     if "d101" in hexdata:
# #         chunks=hexdata.split("d101")
# #         for ts in chunks:
# #             tstamplistd101.append(ts[-12:]+"d101")
            
# #     else:
# #         print "no d101 times at all"
# #     if "d102" in hexdata:
# #         dunks=hexdata.split("d102")
# #         for st in dunks:
# #             tstamplistd101.append(st[-12:]+"d102")
# #     else:
# #         print "no d102 times here mate"
# #     print tstamplistd101
# #     return tstamplistd101
    

# # #In = list of 16 char timestamps in BE order. Out= call Win32Time function supplied with list in LE order 
# # def sortTimeStamps(tslist):
# #     if tslist:
# #         for ts in tslist:
# #             infoMeta.append(ts)
# #             hightime = ts[6:8]+ts[4:6]+ts[2:4]+ts[0:2]
# #             lowtime =ts[14:16]+ts[12:14]+ts[10:12]+ts[8:10]
# #             ft= hightime+":"+lowtime
# #             infoMeta.append(ft)
# #             try:
# #                 win32time(ft)
# #             except:
# #                 print "win32 is probably in all likelihood out of range!"
# #     else:
# #         print "no tslist"


# # def win32time(ft):
# #     #switch parts
# #     h2, h1 = [int(h, base=16) for h in ft.split(':')]
# #     # rebuild
# #     ft_dec = struct.unpack('>Q', struct.pack('>LL', h1, h2))[0]
# #     # add ft_dec to infometa
# #     infoMeta.append(ft_dec)
# #     # use function from iceaway's comment
# #     infoMeta.append(filetime_to_dt(ft_dec))


# # # to do get the main stuff a workin'



# # def Main(url):
# #     infoMeta =[]
# #     print "Really about to dl: ", url
# #     file_name, hexdata = Downloadfile(url)
# #     tstamplistd101 = timestamp(hexdata)
# #     ft = sortTimeStamps(tstamplistd101)
# #     infoMeta.append("\n")
# #     aline=[]
# #     for item in infoMeta:
# #         aline.append(item)
# #     with open("mydocs.csv", "w") as f:
# #         an_entry = csv.writer(f)
# #         an_entry.writerow(aline)
# #     f.close()


# # #url = "https://guccifer2.files.wordpress.com/2016/07/hsu-contributions.xls"
# # #Main(url)

# # with open("attachments.csv", "r") as atts:
# #     attline = csv.reader(atts)

# #     for line in attline:
# #         if "doc" in line[1] and "docx" not in line[1]:
# #             print "Downloading: ", line[1]
# #             Main(line[1])
# #         else:
# #             pass
# #     atts.close()


    

