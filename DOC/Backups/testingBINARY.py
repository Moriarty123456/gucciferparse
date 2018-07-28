import sys


sys.path.append("/home/david/Documents/ClintonJail/Guccifer2.0/Docs&Tools/Python/ALL_GUC_FILE_PROJECT/python/modules/DOC")

from OFFICIALPOSITIONS import *
from ALTERNATEROOTS import *
from ALLFEFF00PARTS import *
#from USERDEFDOCSUM import *
#from rootPARSE import *
#from summaryinfoPARSE import *
#from docsuminfoPARSE import *

def getLocations(infile):
    # if off and altpositions don't exist provide an empty list of same size to aid the csv later on...
    emptylist = ["","","","","","","","","","","","","",""]


    try:
        offPositions = OfficialPositions(infile)
    except:
        offPositions = emptylist


    try:
        altPositions = AlternatePositions(infile)
    except:
        altPositions = emptylist



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


    try:
        locations["OfficialSummaryInfoLocation"] = offPositions[10][2]
    except:
        locations["OfficialSummaryInfoLocation"] = 0
    try:
        locations["OfficialSummaryInfoChunk"] = offPositions[10][4]
    except:
        locations["OfficialSummaryInfoChunk"] = ""
    try:
        locations["OfficialDocSummaryLocation"] = offPositions[6][2]
    except:
        locations["OfficialDocSummaryLocation"] = 0
    try:
        locations["OfficialDocSummaryChunk"] = offPositions[6][4]
    except:
        locations["OfficialDocSummaryChunk"] = ""
    try:  
        locations["OfficialRootChunk"] = offPositions[5]
    except:
        locations["OfficialRootChunk"] = ""
    try:
        locations["OfficialRootTimes"] = offPositions[3]
    except:
        locations["OfficialRootTimes"] = []
    try:
        locations["OfficialFileName"] = offPositions[0]
    except:
        locations["OfficialFileName"] = ""


    try:
        locations["AlternateSummaryInfoLocation"] = altPositions[10][2]
    except:
        locations["AlternateSummaryInfoLocation"] = 0
    try:
        locations["AlternateSummaryInfoChunk"] = altPositions[10][4]
    except:
        locations["AlternateSummaryInfoChunk"] = ""
    try:
        locations["AlternateDocSummaryLocation"] = altPositions[6][2]
    except:
        locations["AlternateDocSummaryLocation"] = 0
    try:
        locations["AlternateDocSummaryChunk"] = altPositions[6][4]
    except:
        locations["AlternateDocSummaryChunk"] = ""
    try:  
        locations["AlternateRootChunk"] = altPositions[5]
    except:
        locations["AlternateRootChunk"] = ""
    try:
        locations["AlternateRootTimes"] = altPositions[3]
    except:
        locations["AlternateRootTimes"] = []
    try:
        locations["AlternateFileName"] = altPositions[0]
    except:
        locations["AlternateFileName"] = ""


    return locations
# # FILE
# print "FILE"
# print "================="
# print locations["OfficialFileName"]

# # OFFICIAL ROOT
# print "OFFICIAL ROOT TIMES"
# print "================="
# print locations["OfficialRootTimes"]

# print "OFFICIAL ROOT CHUNK"
# print "================="
# binarydict = rootPARSE(locations["OfficialRootChunk"])
# print binarydict

# # OFFICIAL SUMMARY INFORMATION
# print "================="
# print "OFFICIAL SUMMARY INFORMATION"
# print "Is there a Summary Info for Orphan?"
# print locations["OfficialSummaryInfoLocation"]
# print locations["AlternateSummaryInfoLocation"]
# print locations["FEFFSumInfoListLocation"]


# print "Summary Info Official"
# print "----------------------------"
# Officialsisdict, Officialsisproplist = simain(locations["OfficialSummaryInfoChunk"])
# print Officialsisdict
# print "========================"
# print Officialsisproplist




# print "========================"
# print "Summary Info For Atlernate Root"
# print "----------------------------"

# Alternatesisdict, Alternatesisproplist = simain(locations["AlternateSummaryInfoChunk"])
# print Alternatesisdict
# print "========================"
# print Alternatesisproplist
# print "========================"

# print "========================"
# print "Offical DOCUMENT SUMMARY"
# print "========================"
# print "orphaned Doc Sum?"
# if locations["OfficialDocSummaryLocation"]:
#     print "There is an off dsl"
#     print dsmain(locations["OfficialDocSummaryChunk"])
# else:
#     print "There is no off dsl"
#     if locations["AlternateDocSummaryLocation"]:
#         print "but there is an Alternate DSL"
#         print dsmain(locations["AlternateDocSummaryChunk"])
#     else:
#         print "nor is there an alt one."
#         if locations["FEFFDocSumListLocation"]:
#             print "But there is an FEFF one."
#             print dsmain(locations["FEFFDocSumListChunk"])
#         else:
#             print "there's no DSL at all"
# if locations["FEFFUserDefListLocation"]:
#     print "there is a single user def location"
#     print locations["FEFFUserDefListLocation"]
# else:
#     print "There's no sinle User DSL."
#     if locations["FEFFDocSumUserDefListLocation"]:
#        print "But there's a UserDef and Doc Sum location"
#        print uddsmain(locations["FEFFDocSumUserDefListChunk"])
