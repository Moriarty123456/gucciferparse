import sys


sys.path.append("/home/david/Documents/ClintonJail/Guccifer2.0/Docs&Tools/Python/ALL_GUC_FILE_PROJECT/python/modules/DOC")
from LOCATIONS import *
from USERDEFDOCSUM import *
from rootPARSE import *
from summaryinfoPARSE import *
from docsuminfoPARSE import *
from FILEINFORMATIONBLOCK import *
from FIB_BLOB import *

######### NOT TESTED ########################################

infile = "/home/david/Documents/ClintonJail/Guccifer2.0/Docs&Tools/Python/ALL_GUC_FILE_PROJECT/downloaded_docs/xls/1st-quarter-targeted-renewals.xls"

#infile = "/home/david/Documents/ClintonJail/Guccifer2.0/Docs&Tools/Python/ALL_GUC_FILE_PROJECT/downloaded_docs/doc/hsu.doc"

#infile = "/home/david/Documents/ClintonJail/Guccifer2.0/Docs&Tools/Python/ALL_GUC_FILE_PROJECT/downloaded_docs/doc/foreign-donations-fpotus-foundations.doc"

#infile = "/home/david/Documents/ClintonJail/Guccifer2.0/Docs&Tools/Python/ALL_GUC_FILE_PROJECT/downloaded_docs/doc/foreign-donations-fpotus-foundations.doc"

#infile = "/home/david/Documents/ClintonJail/Guccifer2.0/Docs&Tools/Python/ALL_GUC_FILE_PROJECT/downloaded_docs/doc/old/briefing-norman-hsu.doc"

###########################
###### NOT WORKING ##########
###########################

#infile = "/home/david/Documents/ClintonJail/Guccifer2.0/Docs&Tools/Python/ALL_GUC_FILE_PROJECT/downloaded_docs/xls/hsu-contributions.xls"


#####################################
##### WORKING ######################
###################################

#infile = "/home/david/Documents/ClintonJail/Guccifer2.0/Docs&Tools/Python/ALL_GUC_FILE_PROJECT/downloaded_docs/xls/out-of-region-ne-donors.xls"


locations = getLocations(infile)

# Make a list that'll be a single row in the csv file
arow = []

# # FILE   
arow.append(locations["OfficialFileName"])

# TIMES
#arow.append(locations["OfficialRootTimes"])

# # ROOT

#BELOW RESULTS IN A COMPUTER SAYS NO MESSAGE if there's no alt or off chunk.
blankRootDict = {'Left Sibling ID': '', 'Next Sector Location & Size': ['', ''], 'Directory Name': '', 'Modified Time': '', 'Creation Time': '', 'Dir Name Length': '', 'Child ID': '', 'State Bits': '', 'Object Type': '', 'Red Black': '', 'CLSID': ['', '', '', ''], 'Right Sibling ID': ''}

if locations["OfficialRootChunk"]:
    binarydict = rootPARSE(locations["OfficialRootChunk"])
    arow.append(binarydict)
else:
    arow.append(blankRootDict)
if locations["AlternateRootChunk"]:
    altBdict = rootPARSE(locations["AlternateRootChunk"])
    arow.append(altBdict)
else:
    arow.append(blankRootDict)

# SUMMARY INFORMATION LOCATIONS
SumInfoLocationsDict = {}

SumInfoLocationsDict["Official Summary Info Location"] = locations["OfficialSummaryInfoLocation"]
SumInfoLocationsDict["Alternate Summary Info Location"] =  locations["AlternateSummaryInfoLocation"]
SumInfoLocationsDict["FEFF Summary Info Location"] =  locations["FEFFSumInfoListLocation"]

arow.append(SumInfoLocationsDict)

# SUMMARY INFORMATION CONTENTS
blanksisdict = {'DWORD-cSections': '', 'NumberPropertiesHex': '', 'FilePointer-SectionSize': '', 'GUID-formatId': '', 'FilePointer-sectionOffset': '', 'WORD-version': '', 'OSMinorVersion': '', 'NumberPropertiesDec': 0, 'OSMajorVersion': '', 'OSType': '', 'WORDbyteOrder': '', 'GUID-applicationClsid': ''}
blanksisproplist = ['GKPIDSI_CODEPAGE', 0, 'GKPIDSI_AUTHOR', '', 'GKPIDSI_LASTAUTHOR', '', 'GKPIDSI_APPNAME', '', 'GKPIDSI_CREATE_DTM', '', 'GKPIDSI_LASTSAVE_DTM', '', 'GKPIDSI_DOC_SECURITY']

if SumInfoLocationsDict["Official Summary Info Location"] == 0:
    Officialsisdict, Officialsisproplist = blanksisdict, blanksisproplist
else:
    Officialsisdict, Officialsisproplist = simain(locations["OfficialSummaryInfoChunk"])
arow.append(Officialsisdict)
arow.append(Officialsisproplist)

if SumInfoLocationsDict["Alternate Summary Info Location"] == 0:
    Alternatesisdict, Alternatesisproplist = blanksisdict, blanksisproplist
else:
    Alternatesisdict, Alternatesisproplist = simain(locations["AlternateSummaryInfoChunk"])
arow.append(Alternatesisdict)
arow.append(Alternatesisproplist)

if SumInfoLocationsDict["FEFF Summary Info Location"] == 0:
    FEFFsisdict, FEFFsisproplist = blanksisdict, blanksisproplist
else:
    FEFFsisdict, FEFFsisproplist = simain(locations["FEFFSumInfoListChunk"])
arow.append(FEFFsisdict)
arow.append(FEFFsisproplist)

# DOCUMENT SUMMARY LOCATIONS

DocInfoLocationsDict = {}

DocInfoLocationsDict["Official Document Summary Info Location"] = locations["OfficialDocSummaryLocation"]
DocInfoLocationsDict["Alternate Document Summary Info Location"] =  locations["AlternateDocSummaryLocation"]
DocInfoLocationsDict["FEFF Document Summary Info Location"] =  locations["FEFFDocSumListLocation"]

arow.append(DocInfoLocationsDict)

# Make blank lists in case there are no off, or alt or FEFF chunks

blankDSList = ['NUMBEROFSECTIONS:', 0, {'WORD-version': '', 'OSMinorVersion': '', 'OSMajorVersion': '', 'OSType': '', 'WORDbyteOrder': '', 'GUID-applicationClsid': ''}, 'DWORD-cSections:', 0, {'SISOffset': '', 'NumberPropertiesDec': 0, 'SISInfoGUID': '', 'SISLengthDec': 0}, ['GKPIDDSI_CODEPAGE', 0, 'GKPIDDSI_COMPANY', '', 'GKPIDDSI_VERSION', 0, 'GKPIDDSI_SCALE', 0, 'GKPIDDSI_LINKSDIRTY', 0, 'GKPIDDSI_SHAREDDOC', 0, 'GKPIDDSI_HYPERLINKSCHANGED', 0, 'GKPIDDSI_DOCPARTS', '', 'GKPIDDSI_HEADINGPAIR']]


blanksDSproplist = ['GKPIDDSI_CODEPAGE', 0, 'GKPIDDSI_COMPANY', '', 'GKPIDDSI_VERSION', 0, 'GKPIDDSI_SCALE', 0, 'GKPIDDSI_LINKSDIRTY', 0, 'GKPIDDSI_SHAREDDOC', 0, 'GKPIDDSI_HYPERLINKSCHANGED', 0, 'GKPIDDSI_DOCPARTS', '', 'GKPIDDSI_HEADINGPAIR']

# add them to a tuple as output
#blankDSMAINtuple =(blankDSList, blanksDSproplist)

if DocInfoLocationsDict["Official Document Summary Info Location"] == 0:
    OfficialDSList = blankDSList
    OfficialDSPropList = blanksDSproplist
else:
    OfficialDSList , OfficialDSPropList = dsmain(locations["OfficialDocSummaryChunk"])

arow.append(OfficialDSList)
arow.append(OfficialDSPropList)

if DocInfoLocationsDict["Alternate Document Summary Info Location"] == 0:
    AlternateDSList = blankDSList
    AlternateDSPropList = blanksDSproplist
else:
    AlternateDSList , AlternateDSPropList = dsmain(locations["AlternateDocSummaryChunk"])

arow.append(AlternateDSList)
arow.append(AlternateDSPropList)

if DocInfoLocationsDict["FEFF Document Summary Info Location"] == 0:
    FEFFDSList = blankDSList
    FEFFDSPropList = blanksDSproplist
else:
    FEFFDSList , FEFFDSPropList = dsmain(locations["FEFFDocSumListChunk"])

arow.append(FEFFDSList)
arow.append(FEFFDSPropList)


#############################################
# DOCUMENT SUMMARY WITH USER DEFINED DICT LOCATIONS
# only in FEFFLocations 
#############################################
DocUSERLocationsDict = {}

DocUSERLocationsDict["Document User Info Location"] =  locations["FEFFDocSumUserDefListLocation"]

arow.append(DocUSERLocationsDict)
#print locations["FEFFDocSumUserDefListChunk"]
blankuserdictlist = [] # TO FILL IN!

if DocUSERLocationsDict["Document User Info Location"] == 0:
    docsumuser = blankuserdictlist
else:
    docsumuser = uddsmain(locations["FEFFDocSumUserDefListChunk"])

arow.append(docsumuser)


#####################################
# USER DEFINED SECTION ON IT'S OWN..?
#####################################

USERLocationsDict = {}

USERLocationsDict["SOLO User Info Location"] =  locations["FEFFUserDefListLocation"]


arow.append(USERLocationsDict)

#########################
# FILEINFORMATION BLOCK #
#########################

filetype = infile.split(".")[1]

if filetype == "doc":
    FIBTitlelist, FIBAnsList =  FIBList(infile)
    arow.append(FIBTitlelist)
    arow.append(FIBAnsList)

print "~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
print arow
