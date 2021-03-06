# In Hsu.doc. we find
# 00 00 20 07  00 00 3E 0C  00 00 5E 13  00 00 3A 01
# 00 00 05 00  12 01 00 00  09 04 00 00  00 00 00 00

# which is a DATE Packet version:
# https://msdn.microsoft.com/en-us/library/dd942025.aspx

# Value (8 bytes): The value of the DATE is an 8-byte IEEE floating-point number, as specified in [MS-OAUT] section 2.2.25.

#      typedef double DATE;

# The date information is represented by whole-number increments, starting with December 30, 1899 midnight as time zero. The time information is represented by the fraction of a day since the preceding midnight. For example, 6:00 A.M. on January 4, 1900 would be represented by the value 5.25 (5 and 1/4 of a day past December 30, 1899).

# from VT_Values.csv
# VT_ARRAY | VT_DATE (0x2007)	MUST be an ArrayHeader followed by a sequence of DATE (Packet Version) packets.

#two = ((float.fromhex("0x3e0c")+float.fromhex("0x5e13"))/365)+1899
#two = ((float.fromhex("0x0c3e0000")+float.fromhex("0x135e0000"))/365)+1899
two = float.fromhex("0x0c3e135e")
three = float.fromhex("0x013a")
print two**-three, two, three
