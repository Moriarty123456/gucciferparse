import sys


sys.path.append("/home/david/Documents/ClintonJail/Guccifer2.0/Docs&Tools/Python/ALL_GUC_FILE_PROJECT/python/modules/DOC")

from OFFICIALPOSITIONS import *
from ALTERNATEROOTS import *
from ALLFEFF00PARTS import *
from USERDEFDOCSUM import *
from rootPARSE import *
from summaryinfoPARSE import *
from docsuminfoPARSE import *

######### NOT TESTED ########################################

#infile = "/home/david/Documents/ClintonJail/Guccifer2.0/Docs&Tools/Python/ALL_GUC_FILE_PROJECT/downloaded_docs/xls/1st-quarter-targeted-renewals.xls"

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

infile = "/home/david/Documents/ClintonJail/Guccifer2.0/Docs&Tools/Python/ALL_GUC_FILE_PROJECT/downloaded_docs/xls/out-of-region-ne-donors.xls"


try:
    offPositions = OfficialPositions(infile)
except:
    pass

try:
    altPositions = AlternatePositions(infile)
except:
    pass

locations = {}

# Check for other relevant Chunks ...
A,B,C,D,E,F = FEFFLists(infile)

if A:
    locations["FEFFSumInfoListLocation"] = A[0]
    locations["FEFFSumInfoListChunk"] = A[1]
else:
    locations["FEFFSumInfoListLocation"] = 0
    locations["FEFFSumInfoListChunk"] = 0
if B:
    locations["FEFFDocSumUserDefListLocation"] = B[0]
    locations["FEFFDocSumUserDefListChunk"] = B[1]
else:
    locations["FEFFDocSumUserDefListLocation"] = 0
    locations["FEFFDocSumUserDefListChunk"] = 0
if C:
    locations["FEFFDocSumListLocation"] = C[0]
    locations["FEFFDocSumListChunk"] = C[1]
else:
    locations["FEFFDocSumListLocation"] = 0
    locations["FEFFDocSumListChunk"] = 0
if D:
    locations["FEFFUserDefListLocation"] = D[0]
    locations["FEFFUserDefListListChunk"] = D[1]
else:
    locations["FEFFUserDefListLocation"] = 0
    locations["FEFFUserDefListChunk"] = 0
if E:
    locations["FEFFAlternateListLocation"] = E[0]
    locations["FEFFAlternateListChunk"] = E[1]
else:
    locations["FEFFAlternateListLocation"] = 0
    locations["FEFFAlternateListChunk"] = 0
if F:
    locations["FEFFPropertyBagListLocation"] = F[0]
    locations["FEFFPropertyBagListChunk"] = F[1]
else:
    locations["FEFFPropertyBagListLocation"] = 0
    locations["FEFFPropertyBagListChunk"] = 0


locations["OfficialSummaryInfoLocation"] = offPositions[10][2]
locations["OfficialSummaryInfoChunk"] = offPositions[10][4]
locations["OfficialDocSummaryLocation"] = offPositions[6][2]
locations["OfficialDocSummaryChunk"] = offPositions[6][4]
locations["OfficialRootChunk"] = offPositions[5]
locations["OfficialRootTimes"] = offPositions[3]
Locations["OfficialFileName"] = offPositions[0]

locations["AlternateSummaryInfoLocation"] = offPositions[10][2]
locations["AlternateSummaryInfoChunk"] = offPositions[10][4]
locations["AlternateDocSummaryLocation"] = offPositions[6][2]
locations["AlternateDocSummaryChunk"] = offPositions[6][4]
locations["AlternateRootChunk"] = offPositions[5]
locations["AlternateRootTimes"] = offPositions[3]
locations["AlternateFileName"] = offPositions[0]
# FILE
print "FILE"
print "================="
print locations["OfficialFileName"]

# OFFICIAL ROOT
print "OFFICIAL ROOT TIMES"
print "================="
print locations["OfficialRootTimes"]

print "OFFICIAL ROOT CHUNK"
print "================="
binarydict = rootPARSE(locations["OfficialRootChunk"])
print binarydict

# OFFICIAL SUMMARY INFORMATION
print "================="
print "OFFICIAL SUMMARY INFORMATION"
print "Is there a Summary Info for Orphan?"
print locations["OfficialSummaryInfoLocation"]
print locations["AlternateSummaryInfoLocation"]
print locations["FEFFSumInfoListLocation"]


print "Summary Info Official"
print "----------------------------"
Officialsisdict, Officialsisproplist = simain(locations["OfficialSummaryInfoChunk"])
print Officialsisdict
print "========================"
print Officialsisproplist




print "========================"
print "Summary Info For Atlernate Root"
print "----------------------------"

Alternatesisdict, Alternatesisproplist = simain(locations["AlternateSummaryInfoChunk"])
print Alternatesisdict
print "========================"
print Alternatesisproplist
print "========================"

print "========================"
print "Offical DOCUMENT SUMMARY"
print "========================"
print "orphaned Doc Sum?"
if locations["OfficialDocSummaryLocation"]:
    print "There is an off dsl"
    print dsmain(locations["OfficialDocSummaryChunk"])
else:
    print "There is no off dsl"
    if locations["AlternateDocSummaryLocation"]:
        print "but there is an Alternate DSL"
        print dsmain(locations["AlternateDocSummaryChunk"])
    else:
        print "nor is there an alt one."
        if locations["FEFFDocSumListLocation"]:
            print "But there is an FEFF one."
            print dsmain(locations["FEFFDocSumListChunk"])
        else:
            print "there's no DSL at all"
if locations["FEFFUserDefListLocation"]:
    print "there is a single user def location"
    print locations["FEFFUserDefListLocation"]
else:
    print "There's no sinle User DSL."
    if locations["FEFFDocSumUserDefListLocation"]:
       print "But there's a UserDef and Doc Sum location"
       print uddsmain(locations["FEFFDocSumUserDefListChunk"])
