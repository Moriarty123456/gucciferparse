# coding: utf-8
#!/usr/bin/env python

import sys

sys.path.append("/home/david/Documents/ClintonJail/Guccifer2.0/Docs&Tools/Python/ALL_GUC_FILE_PROJECT/python/modules/DOC")

from binarytoolkit import *

#infile = "/home/david/Documents/ClintonJail/Guccifer2.0/Docs&Tools/Python/ALL_GUC_FILE_PROJECT/downloaded_docs/xls/1st-quarter-targeted-renewals.xls"

infile = "/home/david/Documents/ClintonJail/Guccifer2.0/Docs&Tools/Python/ALL_GUC_FILE_PROJECT/downloaded_docs/doc/old/foreign-donations-fpotus-foundations.doc"

def FIBBASE(FIBbase):

    FIBbasedict = {}
    FIBbasedict["wIdent"] = FIBbase[0][:4]
    #0 eca5
    #An unsigned integer that specifies that this is a Word Binary File. This value MUST be 0xA5EC.
    nFIB = FIBbase[0][4:]
    # 0 c100
    #nFib (2 bytes): An unsigned integer that specifies the version number of the file format used. Superseded by FibRgCswNew.nFibNew if it is present. This value SHOULD<12> be 0x00C1.
    FIBbasedict["unused"] = FIBbase[1][:4]
    # 1 59e0

    FIBbasedict["lid"] = FIBbase[1][4:]
    #1 0904

    # (2 bytes): A LID that specifies the install language of the application that is producing the document. If nFib is 0x00D9 or greater, then any East Asian install lid or any install lid with a base language of Spanish, German or French MUST be recorded as lidAmerican. If the nFib is 0x0101 or greater, then any install lid with a base language of Vietnamese, Thai, or Hindi MUST be recorded as lidAmerican.
    FIBbasedict["pnNext"] = FIBbase[2][:4]
    # 2 0000
    #pnNext (2 bytes): An unsigned integer that specifies the offset in the WordDocument stream of the FIB for the document which contains all the AutoText items. If this value is 0, there are no AutoText items attached. Otherwise the FIB is found at file location pnNext×512. If fGlsy is 1 or fDot is 0, this value MUST be 0. If pnNext is not 0, each FIB MUST share the same values for FibRgFcLcb97.fcPlcBteChpx, FibRgFcLcb97.lcbPlcBteChpx, FibRgFcLcb97.fcPlcBtePapx, FibRgFcLcb97.lcbPlcBtePapx, and FibRgLw97.cbMac.
    binsection1 = hextoBINARYlist(FIBbase[2][4],16)
    FIBbasedict["fDot"] = binsection1[0]
    #2 f
    # A - fDot (1 bit): Specifies whether this is a document template.
    FIBbasedict["fGlsy"] = binsection1[1]
    #2 0
    #B - fGlsy (1 bit): Specifies whether this is a document that contains only AutoText items (see FibRgFcLcb97.fcSttbfGlsy, FibRgFcLcb97.fcPlcfGlsy and FibRgFcLcb97.fcSttbGlsyStyle).
    FIBbasedict["fComplex"] = binsection1[2]
    #2 1
    #C - fComplex (1 bit): Specifies that the last save operation that was performed on this document was an incremental save operation.
    FIBbasedict["fHasPic"] = binsection1[3]
    #2 2
    #D - fHasPic (1 bit): When set to 0, there SHOULD<13> be no pictures in the document.
    FIBbasedict["cQuickSaves"] = int(str("".join(binsection1[4:8])),16)
    # 3 bf00
    #E - cQuickSaves (4 bits): An unsigned integer. If nFib is less than 0x00D9, then cQuickSaves specifies the number of consecutive times this document was incrementally saved. If nFib is 0x00D9 or greater, then cQuickSaves MUST be 0xF.
    FIBbasedict["fEncrypted"] = binsection1[8]
    #3 0
    # F - fEncrypted (1 bit): Specifies whether the document is encrypted or obfuscated as specified in Encryption and Obfuscation.
    FIBbasedict["fWhichTblStm"] = binsection1[9]
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
    FIBbasedict["fLoadOverride"] = binsection1[13]
    # 4 0
    #K - fLoadOverride (1 bit): Specifies whether to override the language information and font that are specified in the paragraph style at istd 0 (the normal style) with the defaults that are appropriate for the installation language of the application.
    FIBbasedict["fFarEast"] = binsection1[14]
    #4 0
    # L - fFarEast (1 bit): Specifies whether the installation language of the application that created the document was an East Asian language.
    FIBbasedict["fObfuscated"] = binsection1[15]
    #M - fObfuscated (1 bit): If fEncrypted is 1, this bit specifies whether the document is obfuscated by using XOR obfuscation (section 2.2.6.1); otherwise, this bit MUST be ignored.
    #4 0
    FIBbasedict["nFibBack"] = FIBbase[3][4:]
    #(2 bytes): This value SHOULD<14> be 0x00BF. This value MUST be 0x00BF or 0x00C1.
    # 4 0010 !!!! 3 = bf000000
    FIBbasedict["lKey"] = FIBbase[4][:4]
    # 5 0000
    # (4 bytes): If fEncrypted is 1 and fObfuscation is 1, this value specifies the XOR obfuscation (section 2.2.6.1) password verifier. If fEncrypted is 1 and fObfuscation is 0, this value specifies the size of the EncryptionHeader that is stored at the beginning of the Table stream as described in Encryption and Obfuscation. Otherwise, this value MUST be 0.
    FIBbasedict["envr"] = FIBbase[4][4:6]
    # 5 0
    #(1 byte): This value MUST be 0, and MUST be ignored.
    binsection2 = hextoBINARYlist(FIBbase[4][6:],8) 

    FIBbasedict["fMac"] = binsection2[0] 
    #5 0
    #     (1 bit): This value MUST be 0, and MUST be ignored.
    FIBbasedict["fEmptySpecial"] = binsection2[1]
    #     (1 bit): This value SHOULD<15> be 0 and SHOULD<16> be ignored.
    FIBbasedict["fLoadOverridePage"] = binsection2[2]
    # 5 0
    #(1 bit): Specifies whether to override the section properties for page size, orientation, and margins with the defaults that are appropriate for the installation language of the application.
    FIBbasedict["reserved1"] = binsection2[3]
    #     (1 bit): This value is undefined and MUST be ignored.
    FIBbasedict["reserved2"] = binsection2[4]
    #     (1 bit): This value is undefined and MUST be ignored.
    FIBbasedict["fSpare0"] = int(str("".join(binsection2[5:])),16)
    #     (3 bits): This value is undefined and MUST be ignored.
    FIBbasedict["reserved3"] = FIBbase[5][:2]
    #     (2 bytes): This value MUST be 0 and MUST be ignored.
    FIBbasedict["reserved4"] = FIBbase[5][2:4]
    #     (2 bytes): This value MUST be 0 and MUST be ignored.
    FIBbasedict["reserved5"] = FIBbase[5][4:]
    #     (4 bytes): This value is undefined and MUST be ignored.
    FIBbasedict["reserved6"] = FIBbase[6][:4]
    #     (4 bytes): This value is undefined and MUST be ignored.
    return nFIB, FIBbasedict



def FileINFORMATIONBLOCK(infile):
    #################
    # SEE MSDOC.PDF #
    # PAGE 52       #
    #################
    FIBDict = {}
    filename = infile.split("/")[-1]
    FIBDict["FILENAME"] = filename
    herx = OutputHex(infile)
    hexchunks = Header2FourBytes(herx)
    # first sector of the FIB.
    FIB1 = hexchunks[128:256]
    FIBbase = FIB1[0:8]

    ##################################
    # The FibBase structure is the   #
    # fixed-size portion of the Fib. #
    # It contains the nFib used for  #
    # calculating other values.      #
    ##################################
    
    nFib, FIBbasedict = FIBBASE(FIBbase)
    FIBDict["nFib"] = nFib
    FIBDict["FIBbasedict"] = FIBbasedict
    FIBDict["Lang Installed on App Creating Doc"] = FIBbasedict["lid"]
    FIBDict["csw"] = FIB1[8][:4] #csw MUST be 0x000E.
    flist = []
    flist.append(FIB1[8][4:])
    for f in FIB1[9:14]:
        flist.append(f)
    flist.append(FIB1[15][:4])
    FIBDict["fibRgW"] = flist # The FibRgW97
    FIBDict["lidFE"] = flist[-1]
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
    FIBDict["cslw"] = FIB1[15][4:] # MUST be 0x0016.
    FIBDict["fibRgLw"] = FIB1[16:38]
    FIBDict["cbMac"] = FIB1[16]
    # 4 bytes :  Specifies the count of bytes of those written to the WordDocument stream of the file that have any meaning. All bytes in the WordDocument stream at offset cbMac and greater MUST be ignored.
    FIBDict["reserved1"] = FIB1[17]
    # 4 bytes :  This value is undefined and MUST be ignored.
    FIBDict["reserved2"] = FIB1[18]
    # 4 bytes :  This value is undefined and MUST be ignored.
    FIBDict["ccpText"] = FIB1[19]
    # 4 bytes :  A signed integer that specifies the count of CPs in the main document. This value MUST be zero 1  or greater.  
    FIBDict["ccpFtn"] = FIB1[20]
    # 4 bytes :  A signed integer that specifies the count of CPs in the footnote subdocument. This value MUST be zero 1  or greater.  
    FIBDict["ccpHdd"] = FIB1[21]
    # 4 bytes :  A signed integer that specifies the count of CPs in the header subdocument. This value MUST be zero 1  or greater.  
    FIBDict["reserved3"] = FIB1[22]
    # 4 bytes :  This value MUST be zero and MUST be ignored.
    FIBDict["ccpAtn"] = FIB1[23]
    # 4 bytes :  A signed integer that specifies the count of CPs in the comment subdocument. This value MUST be zero 1  or greater.  
    FIBDict["ccpEdn"] = FIB1[24]
    # 4 bytes :  A signed integer that specifies the count of CPs in the endnote subdocument. This value MUST be zero 1  or greater.  
    FIBDict["ccpTxbx"] = FIB1[25]
    # 4 bytes :  A signed integer that specifies the count of CPs in the textbox subdocument of the main document. This value MUST be zero 1  or greater.  
    FIBDict["ccpHdrTxbx"] = FIB1[26]
    # 4 bytes :  A signed integer that specifies the count of CPs in the textbox subdocument of the header. This value MUST be zero 1  or greater.  
    FIBDict["reserved4"] = FIB1[27]
    # 4 bytes :  This value is undefined and MUST be ignored.
    FIBDict["reserved5"] = FIB1[28]
    # 4 bytes :  This value is undefined and MUST be ignored.
    FIBDict["reserved6"] = FIB1[29]
    # 4 bytes :  This value MUST be equal or less than the number of data elements in PlcBteChpx  as specified by FibRgFcLcb97.fcPlcfBteChpx and FibRgFcLcb97.lcbPlcfBteChpx. This value MUST be ignored. 
    FIBDict["reserved7"] = FIB1[30]
    # 4 bytes :  This value is undefined and MUST be ignored
    FIBDict["reserved8"] = FIB1[31]
    # 4 bytes :  This value is undefined and MUST be ignored
    FIBDict["reserved9"] = FIB1[32]
    # 4 bytes :  This value MUST be less than or equal to the number of data elements in PlcBtePapx  as specified by FibRgFcLcb97.fcPlcfBtePapx andFibRgFcLcb97.lcbPlcfBtePapx. This value MUST be ignored. 
    FIBDict["reserved10"] = FIB1[33]
    # 4 bytes :  This value is undefined and MUST be ignored.
    FIBDict["reserved11"] = FIB1[34]
    # 4 bytes :  This value is undefined and MUST be ignored.
    FIBDict["reserved12"] = FIB1[35]
    # 4 bytes :  This value SHOULD <26> be zero  and MUST be ignored. 
    FIBDict["reserved13"] = FIB1[36]
    # 4 bytes :  This value MUST be zero and MUST be ignored.
    FIBDict["reserved14"] = FIB1[37]
    # 4 bytes :  This value MUST be zero and MUST be ignored.
    FIBDict["cbRgFcLcb"] = FIB1[38][:4] # !! Gives b700 (183), with nFib c100!
    ##################################################
    # int:count of 64-bit values  in fibRgFcLcbBlob. #
    # MUST be one of the following values,           #
    # depending on the value of nFib.                #
    # Value of nFib cbRgFcLcb                        #
    # 0x00C1 0x005D                                  #
    # 0x00D9 0x006C                                  #
    # 0x0101 0x0088                                  #
    # 0x010C 0x00A4                                  #
    # 0x0112 0x00B7                                  #
    ##################################################
    y = 39
    for x in FIB1[39:224]:
        print y, ",", x
        print ""
        y = y+1


    fibRgFcLcbBlob = "" #(variable): The FibRgFcLcb.
    cswNew = ""
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
    fibRgCswNew = ""
    ##########################################################
    # (variable): If cswNew is nonzero, this is fibRgCswNew. #
    # Otherwise, it is not present in the file.              #
    ##########################################################
    
    for n in FIBDict:
        print n, FIBDict[n]
        #, csw, fibRgW, cslw, fibRgLw, cbRgFcLcb

FileINFORMATIONBLOCK(infile)
