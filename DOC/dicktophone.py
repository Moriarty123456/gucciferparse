import sys


sys.path.append("/home/david/Documents/ClintonJail/Guccifer2.0/Docs&Tools/Python/ALL_GUC_FILE_PROJECT/python/modules/DOC")

from binarytoolkit import *


def dicktophone(dick):
    #input a dictionary - the bit given by Dictionary identifier 00000000
    # eg dick = ['05000000', '02000000', '14000000', '5f004100', '64004800', '6f006300', '52006500', '76006900', '65007700', '43007900', '63006c00', '65004900', '44000000 ...
    #output the text of those fields with number
    # eg {2: '_AdHocReviewCycleID', 3: '_NewReviewCycle', 4: '_EmailSubject', 5: '_AuthorEmail', 6: '_AuthorEmailDisplayName'}
    numof = int(ReverseHex(dick[0]),16)
    mydick = {}
    #posn is counter of position
    posn = 2
    startnum = int(ReverseHex(dick[posn-1]),16)
    #startlength = (int(ReverseHex(dick[posn+1]),16))/2
    #print startlength
    itemno = 0
    for x in range(numof):
        itemno = x + startnum
        #  print "item no is ", itemno,
        # hack to solve problem about modulus
        t = (int(ReverseHex(dick[posn]),16))
        if t/2 == (t-1)/2:
            length = (t+1)/2
        else:
            length = t/2
            #    print "length for ", dick[posn], "is ", length
            #   print "result ", dick[posn+1:posn+1+length]

        thetext = "".join(dick[posn+1:posn+1+length]).decode("hex")
        if "\x00" in thetext:
            thetext = "".join(thetext.split("\x00"))
        else:
            pass
        mydick[itemno] = thetext
        posn = posn + (length+2)
    return mydick


