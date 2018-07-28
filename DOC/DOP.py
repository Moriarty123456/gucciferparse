# coding: utf-8
#!/usr/bin/env python

import sys

sys.path.append("/home/david/Documents/ClintonJail/Guccifer2.0/Docs&Tools/Python/ALL_GUC_FILE_PROJECT/python/modules/DOC")


#######################################
# Page 145 MS-DOC                     #
# if cswNew == 0:                     #
#     thedop = Dop97                  #
# elif FibRgCswNew.nFibNew == "00d9": #
#     thedop = Dop2000                #
# elif FibRgCswNew.nFibNew == "0101": #
#     thedop = Dop2002                #
# elif FibRgCswNew.nFibNew == "010c": #
#     thedop = Dop2003                #
# elif FibRgCswNew.nFibNew == "0112": #
#     thedop = Dop2007                #
# else:                               #
#     thedop = Dop97                  #
#######################################

#dop start 0x4520

def DOPbase(dophex, cswNew):
    dopbasedict = {}
    
