# see here:

# https://msdn.microsoft.com/en-us/library/dd906793(v=office.12).aspx

# The BOF
# vers (2 bytes): An unsigned integer that specifies the BIFF version of the file. The value MUST be 0x0600.

# dt (2 bytes): An unsigned integer that specifies the document type of the substream of records following this record. For more information about the layout of the sub-streams in the workbook stream see File Structure. MUST be a value from the following table:

# Value, Meaningm 0x0005, Specifies the workbook substream.

#     0x0010,  Specifies the dialog sheet substream or the worksheet substream.

# Then look at offset 0x200 in all xls files,

# 00000200   09 08 10 00  00 06 05 00  A0 19 CD 07  C9 C0 00 00  ................
# 00000210   06 03 00 00  E1 00 02 00  B0 04 C1 00  02 00 00 00  ................
# 00000220   E2 00 00 00  5C 00 70 00  07 00 00 55  53 45 52 2D  ....\.p....USER-
# 00000230   50 43 46 6F  72 63 65 20  20 20 20 20  20 20 20 20  PCForce
# 00000240   20 20 20 20  20 20 20 20  20 20 20 20  20 20 20 20
# 00000250   20 20 20 20  20 20 20 20  20 20 20 20  20 20 20 20
# 00000260   20 20 20 20  20 20 20 20  20 20 20 20  20 20 20 20
# 00000270   20 20 20 20  20 20 20 20  20 20 20 20  20 20 20 20
# 00000280   20 20 20 20  20 20 20 20  20 20 20 20  20 20 20 20
# 00000290   20 20 20 20  20 20 20 20  42 00 02 00  B0 04 61 01          B.....a.

