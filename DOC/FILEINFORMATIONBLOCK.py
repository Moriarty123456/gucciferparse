# coding: utf-8
#!/usr/bin/env python

import sys

sys.path.append("/home/david/Documents/ClintonJail/Guccifer2.0/Docs&Tools/Python/ALL_GUC_FILE_PROJECT/python/modules/DOC")

from binarytoolkit import *
from FIB_BLOB import *


def FIBBASE(FIBbase):

    FIBbasedict = {}
    wIdent = FIBbase[0][:4]
    intwIdent = ReverseInt(wIdent)
    if wIdent == "eca5":
        FIBbasedict["FileType"] = "WORD Binary File"
        FIBbasedict["FileTypeNumber"] = intwIdent
    elif wIdent == "0908":
        FIBbasedict["FileType"] = "EXCEL Binary File"
        FIBbasedict["FileTypeNumber"] = intwIdent
    else:
        FIBbasedict["FileType"] = "File Type Unknown"
        FIBbasedict["FileTypeNumber"] = intwIdent
        #0 eca5
    #An unsigned integer that specifies that this is a Word Binary File. This value MUST be 0xA5EC.
    nFIB = FIBbase[0][4:]
    FIBbasedict["nFIB : File Type Version"] = ReverseInt(nFIB)
    # 0 c100
    #nFib (2 bytes): An unsigned integer that specifies the version number of the file format used. Superseded by FibRgCswNew.nFibNew if it is present. This value SHOULD<12> be 0x00C1.
    #FIBbasedict["unused"] = FIBbase[1][:4]
    # 1 59e0
    lid = FIBbase[1][4:]
    FIBbasedict["Language ID of Application"] = ReverseInt(FIBbase[1][4:])
    #1 0904

    # (2 bytes): A LID that specifies the install language of the application that is producing the document. If nFib is 0x00D9 or greater, then any East Asian install lid or any install lid with a base language of Spanish, German or French MUST be recorded as lidAmerican. If the nFib is 0x0101 or greater, then any install lid with a base language of Vietnamese, Thai, or Hindi MUST be recorded as lidAmerican.
    FIBbasedict["pnNext"] = FIBbase[2][:4]
    # 2 0000
    #pnNext (2 bytes): An unsigned integer that specifies the offset in the WordDocument stream of the FIB for the document which contains all the AutoText items. If this value is 0, there are no AutoText items attached. Otherwise the FIB is found at file location pnNext×512. If fGlsy is 1 or fDot is 0, this value MUST be 0. If pnNext is not 0, each FIB MUST share the same values for FibRgFcLcb97.fcPlcBteChpx, FibRgFcLcb97.lcbPlcBteChpx, FibRgFcLcb97.fcPlcBtePapx, FibRgFcLcb97.lcbPlcBtePapx, and FibRgLw97.cbMac.
    binsection1 = hextoBINARYlist(FIBbase[2][4:],16)
    FIBbasedict["fDot Doc Template?"] = binsection1[0]
    #2 f
    # A - fDot (1 bit): Specifies whether this is a document template.
    FIBbasedict["fGlsy Only Autotext?"] = binsection1[1]
    #2 0
    #B - fGlsy (1 bit): Specifies whether this is a document that contains only AutoText items (see FibRgFcLcb97.fcSttbfGlsy, FibRgFcLcb97.fcPlcfGlsy and FibRgFcLcb97.fcSttbGlsyStyle).
    FIBbasedict["fComplex Last save was Autosave?"] = binsection1[2]
    #2 1
    #C - fComplex (1 bit): Specifies that the last save operation that was performed on this document was an incremental save operation.
    FIBbasedict["fHasPic"] = binsection1[3]
    #2 2
    #D - fHasPic (1 bit): When set to 0, there SHOULD<13> be no pictures in the document.
    FIBbasedict["cQuickSaves Num Quicksaves"] = int(str("".join(binsection1[4:8])),16)
    # 3 bf00
    #E - cQuickSaves (4 bits): An unsigned integer. If nFib is less than 0x00D9, then cQuickSaves specifies the number of consecutive times this document was incrementally saved. If nFib is 0x00D9 or greater, then cQuickSaves MUST be 0xF.
    FIBbasedict["fEncrypted"] = binsection1[8]
    #3 0
    # F - fEncrypted (1 bit): Specifies whether the document is encrypted or obfuscated as specified in Encryption and Obfuscation.
    FIBbasedict["fWhichTblStm TableStream 1 or 0?"] = binsection1[9]
    #3 0
    #G - fWhichTblStm (1 bit): Specifies the Table stream to which the FIB refers. When this value is set to 1, use 1Table; when this value is set to 0, use 0Table.
    FIBbasedict["fReadOnlyRecommended"] = binsection1[10]
    # 3 0
    # H - fReadOnlyRecommended (1 bit): Specifies whether the document author recommended that the document be opened in read-only mode.
    FIBbasedict["fWriteReservation"] = binsection1[11]
    # 3 0
    #I - fWriteReservation (1 bit): Specifies whether the document has a write-reservation password
    FIBbasedict["fExtChar"] = binsection1[12]
    # 4 0
    #J - fExtChar (1 bit): This value MUST be 1.
    FIBbasedict["fLoadOverride Language?"] = binsection1[13]
    # 4 0
    #K - fLoadOverride (1 bit): Specifies whether to override the language information and font that are specified in the paragraph style at istd 0 (the normal style) with the defaults that are appropriate for the installation language of the application.
    FIBbasedict["fFarEast Lang Install Far East?"] = binsection1[14]
    #4 0
    # L - fFarEast (1 bit): Specifies whether the installation language of the application that created the document was an East Asian language.
    FIBbasedict["fObfuscated"] = binsection1[15]
    #M - fObfuscated (1 bit): If fEncrypted is 1, this bit specifies whether the document is obfuscated by using XOR obfuscation (section 2.2.6.1); otherwise, this bit MUST be ignored.
    #4 0
    FIBbasedict["nFibBack"] = FIBbase[3][:4]
    #(2 bytes): This value SHOULD<14> be 0x00BF. This value MUST be 0x00BF or 0x00C1.
    # BF00
    FIBbasedict["lKey Password?"] = FIBbase[3][4:]+FIBbase[4][:4]
    # 5 0000 + 0000
    # (4 bytes): If fEncrypted is 1 and fObfuscation is 1, this value specifies the XOR obfuscation (section 2.2.6.1) password verifier. If fEncrypted is 1 and fObfuscation is 0, this value specifies the size of the EncryptionHeader that is stored at the beginning of the Table stream as described in Encryption and Obfuscation. Otherwise, this value MUST be 0.
    #FIBbasedict["envr"] = FIBbase[4][4:6]
    #  00
    #(1 byte): This value MUST be 0, and MUST be ignored.
    binsection2 = hextoBINARYlist(FIBbase[4][6:],8) 

    FIBbasedict["fMac"] = binsection2[0] 
    #5 0
    #     (1 bit): This value MUST be 0, and MUST be ignored.
    #FIBbasedict["fEmptySpecial"] = binsection2[1]
    #     (1 bit): This value SHOULD<15> be 0 and SHOULD<16> be ignored.
    FIBbasedict["fLoadOverridePage text orientation?"] = binsection2[2]
    # 5 0
    #(1 bit): Specifies whether to override the section properties for page size, orientation, and margins with the defaults that are appropriate for the installation language of the application.
    #FIBbasedict["reserved1"] = binsection2[3]
    #     (1 bit): This value is undefined and MUST be ignored.
    #FIBbasedict["reserved2"] = binsection2[4]
    #     (1 bit): This value is undefined and MUST be ignored.
    #FIBbasedict["fSpare0"] = int(str("".join(binsection2[5:])),16)
    #     (3 bits): This value is undefined and MUST be ignored.
    #FIBbasedict["reserved3"] = FIBbase[5][:4]
    #     (2 bytes): This value MUST be 0 and MUST be ignored.#0000
    #FIBbasedict["reserved4"] = FIBbase[5][4:]
    #     (2 bytes): This value MUST be 0 and MUST be ignored.
    #FIBbasedict["reserved5"] = FIBbase[6]
    #     (4 bytes): This value is undefined and MUST be ignored. 00 08 00 00
    #FIBbasedict["reserved6"] = FIBbase[7]
    #     (4 bytes): This value is undefined and MUST be ignored. d1 13 00 00
    return nFIB, FIBbasedict



def FileINFORMATIONBLOCK(infile):
    #################
    # SEE MSDOC.PDF #
    # PAGE 52       #
    #################
    ################################################
    # The Fib is a variable length structure.      #
    # Starting at sector zero i.e. offset          #
    # 0x200                                        #
    # With the exception of the base portion       #
    # which is fixed in size,                      #
    # every section is preceded with a count field #
    # that specifies the size of the next section. #
    # These are                                    #
    # * csw for the The FibRgW97                   #
    # * cslw for the FibRgLw97.                    #
    # and cbRgFcLcb for the fibRgFcLcbBlob         #
    ################################################
    FIBDict = {}
    filename = infile.split("/")[-1]
    FIBDict["FILENAME"] = filename
    herx = OutputHex(infile)
    #hexchunks = Header2FourBytes(herx)
    # first sector of the FIB.
    #FIB1 = hexchunks[128:256]
    #FIBbase = FIB1[0:8]
    FIBbase = Header2FourBytes(herx[1024:1088])

    ##################################
    # The FibBase structure is the   #
    # fixed-size portion of the Fib. #
    # It contains the nFib used for  #
    # calculating other values.      #
    ##################################
    
    nFib, FIBbasedict = FIBBASE(FIBbase)
    FIBDict["nFib"] = nFib
    FIBDict["FIBbasedict"] = FIBbasedict
    #FIBDict["csw"] = ReverseInt(FIB1[8][:4])
    FIBDict["csw"] = ReverseInt(herx[1088:1092])
    #csw An unsigned integer that specifies the count of 16-bit (2 Byte) values corresponding to fibRgW that follow. MUST be 0x000E = 14 x 2 Bytes = 28 Bytes  = 7 x 4byte chunks
    
    # fibRgW all reserved values except last one: the lidFE

    #FIBDict["fibRgW"] = Header2FourBytes(herx[1092:1148])
    # FIBDict["lidFE"] = flist[-1]
    FIBDict["lidFE Language of Stored Style Name"] = Header2FourBytes(herx[1092:1148])[-1]

    # lidFE (2 bytes): A LID whose meaning depends on the nFib value, which is one of the following.
    ###########################################################
    # nFib value Meaning                                      #
    # 0x00C1 If FibBase.fFarEast is "true",                   #
    #        this is the LID of the stored style names.       #
    #        Otherwise it MUST be ignored.                    #
    # 0x00D9                                                  #
    # 0x0101                                                  #
    # 0x010C                                                  #
    # 0x0112 The LID of the stored style names (STD.xstzName) #
    ###########################################################
    cslw = herx[1148:1152] # FIB1[15][4:]
    # An unsigned integer that specifies the count of 32-bit
    # values corresponding to fibRgLw that follow.
    # 32 bit = 4 bytes
    # MUST be 0x0016 == 22 dec x 4 bytes = 88 bytes = 22 chunks
    #FIBDict["fibRgLw"] = Header2FourBytes(herx[1152:1328]) # FIBRGLW(FIB1)
    # OMIT fibRgLw AS IT IS USELESS TO US
    
    cbRgFcLcb = herx[1328:1332] # FIB1[38][:4]
    # 00000298   B7 00 00 00
    ##########################################################
    # cbRgFcLcb                                              #
    # (2 bytes): An unsigned integer that specifies the      #
    # count of 64-bit values corresponding to fibRgFcLcbBlob #
    # that follow. This MUST be one of the following values, #
    # depending on the value of nFib.                        #
    # Value of nFib cbRgFcLcb                                #
    # 0x00C1 0x005D                                          #
    # 0x00D9 0x006C                                          #
    # 0x0101 0x0088                                          #
    # 0x010C 0x00A4                                          #
    # 0x0112 0x00B7                                          #
    ##########################################################
    #bloblength = ReverseInt(FIBDict["cbRgFcLcb"]) # --> 183
    # Therefore 183 x 64-bits = 183 x 8 bytes
    # = 1464 bytes
    startBLOB = 1332
    endBLOB = (2*(ReverseInt(cbRgFcLcb)*8))+1332
    fibRgFcLcbBlob  =  Header2FourBytes(herx[startBLOB:endBLOB])
    #FIBDict["fibRgFcLcbBlob"] = theBLOB(fibRgFcLcbBlob) #herx, bloblength)
    if cbRgFcLcb == "5d00":
        FIBDict["fibRgFcLcb97"] = FibRgFcLcb97(fibRgFcLcbBlob[0:186])
        fibRgFcLcb2000 = ""
        fibRgFcLcb2002 = ""
        fibRgFcLcb2003 = "" 
        fibRgFcLcb2007 = ""
    elif cbRgFcLcb == "6c00":
        FIBDict["fibRgFcLcb97"] = FibRgFcLcb97(fibRgFcLcbBlob[0:186])
        FIBDict["fibRgFcLcb2000"] = FibRgFcLcb2000(fibRgFcLcbBlob[186:216])
        fibRgFcLcb2002 = ""
        fibRgFcLcb2003 = ""
        fibRgFcLcb2007 = ""
    elif cbRgFcLcb =="8800":
        FIBDict["fibRgFcLcb97"] = FibRgFcLcb97(fibRgFcLcbBlob[0:186])
        FIBDict["fibRgFcLcb2000"] = FibRgFcLcb2000(fibRgFcLcbBlob[186:216])
        FIBDict["fibRgFcLcb2002"] = FibRgFcLcb2002(fibRgFcLcbBlob[216:272])
        fibRgFcLcb2003 = ""
        fibRgFcLcb2007 = ""
    elif cbRgFcLcb == "a400":
        FIBDict["fibRgFcLcb97"] = FibRgFcLcb97(fibRgFcLcbBlob[0:186])
        FIBDict["fibRgFcLcb2000"] = FibRgFcLcb2000(fibRgFcLcbBlob[186:216])
        FIBDict["fibRgFcLcb2002"] = FibRgFcLcb2002(fibRgFcLcbBlob[216:272])
        FIBDict["fibRgFcLcb2003"] = FibRgFcLcb2003(fibRgFcLcbBlob[272:328])
        fibRgFcLcb2007 = ""
    elif cbRgFcLcb == "b700":
        FIBDict["fibRgFcLcb97"] = FibRgFcLcb97(fibRgFcLcbBlob[0:186])
        FIBDict["fibRgFcLcb2000"] = FibRgFcLcb2000(fibRgFcLcbBlob[186:216])
        FIBDict["fibRgFcLcb2002"] = FibRgFcLcb2002(fibRgFcLcbBlob[216:272])
        FIBDict["fibRgFcLcb2003"] = FibRgFcLcb2003(fibRgFcLcbBlob[272:328])
        FIBDict["fibRgFcLcb2007"] = FibRgFcLcb2007(fibRgFcLcbBlob[328:len(fibRgFcLcbBlob)])
    else:
        FIBDict["fibRgFcLcb97"] = ""
        FIBDict["fibRgFcLcb2000"] = ""
        FIBDict["fibRgFcLcb2002"] = ""
        FIBDict["fibRgFcLcb2003"] = ""
        FIBDict["fibRgFcLcb2007"] = ""

    # FibRgFcLcb97 contains a timestamp of last saved
    try:
        atimestamp = [FIBDict["fibRgFcLcb97"]["dwLowDateTime"],FIBDict["fibRgFcLcb97"]["dwHighDateTime"]]
        FIBDict["Last Saved"] =  TimeStamp(atimestamp)
    except KeyError:
        FIBDict["Last Saved"] =  ""
    
    cswNew = herx[endBLOB:endBLOB+4]
    ##############################################
    # cswNew (2 bytes): An unsigned integer that #
    # specifies the count of 16-bit values       #
    # corresponding to fibRgCswNew that follow.  #
    # This MUST be one of the following values,  #
    # depending on the value of nFib.            #
    # Value of nFib cswNew                       #
    # 0x00C1 0                                   #
    # 0x00D9 0x0002                              #
    # 0x0101 0x0002                              #
    # 0x010C 0x0002                              #
    # 0x0112 0x0005                              #
    ##############################################
    fibRgCswNew = []
    CSWSTART = endBLOB+4
    count =0
    for x in range(ReverseInt(cswNew)):
        fibRgCswNew.append(herx[CSWSTART+count:CSWSTART+4+count])
        count = count +4

    nFibNew = fibRgCswNew[0]
    FIBDict["Version Number FileFormat"] = ReverseInt(nFibNew)
    # An unsigned integer that specifies the version number of the file format that is used. This value MUST be one of the following.
    # Value
    # 0x00D9
    # 0x0101
    # 0x010C
    # 0x0112 
    if nFibNew == "1201":
        FIBDict["cQuickSavesNew"] = fibRgCswNew[1]
        FIBDict["lidThemeOther"] = fibRgCswNew[2]
        FIBDict["lidThemeFE"] = fibRgCswNew[3]
        FIBDict["lidThemeCS"] = fibRgCswNew[4]
    elif nFibNew == "0c01" or nFibNew == "0101" or nFibNew == "d900":
        FIBDict["cQuickSavesNew"] = fibRgCswNew[1]
        FIBDict["lidThemeOther"] = ""
        FIBDict["lidThemeFE"] = ""
        FIBDict["lidThemeCS"] = ""
    else:
        FIBDict["cQuickSavesNew"] = ""
        FIBDict["lidThemeOther"] = ""
        FIBDict["lidThemeFE"] = ""
        FIBDict["lidThemeCS"] = ""
    # note: p100 [MS-DOC] --> This value is undefined and MUST be ignored.
    return FIBDict


def FIBList(infile):
    FinalDict = FileINFORMATIONBLOCK(infile)

    FIBTitlelist = []
    FIBAnsList=[]

    big_list = ["FILENAME" , "nFib" , "Last Saved" , "Version Number FileFormat" , "lidFE Language of Stored Style Name" , "lidThemeFE" , "lidThemeOther" , "lidThemeCS" , "cQuickSavesNew" , "csw"]

    for l in big_list:
        #print l, FinalDict[l]
        FIBTitlelist.append(l)
        FIBAnsList.append(FinalDict[l])

    other_list = [ "FIBbasedict", "fibRgFcLcb97", "fibRgFcLcb2000" , "fibRgFcLcb2002" , "fibRgFcLcb2003" ,  "fibRgFcLcb2007"]
    #print FinalDict

    for x in range(len(other_list)):
        inddict = other_list[x]
        for y in FinalDict[inddict]:
            #print inddict+"_"+y, FinalDict[inddict][y]
            FIBTitlelist.append(inddict+"_"+y)
            FIBAnsList.append(FinalDict[inddict][y])
    return FIBTitlelist, FIBAnsList
