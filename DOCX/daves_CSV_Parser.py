import csv

# INPUT: a) the row, b) a string to search for in row, c) divider (usually same as the string), and d) the end of the string.
# EXAMPLE: To extract email from perkins coie "t>dbaur@perkinscoie.com</t"
# SET VARIABLES AS: a) a row from a CSV file, b) "perkinscoie.com</t" c) ">", d) "</t"
# OUTPUT: List with matches
def row_multiple(row, searchitem, divider, end):
    arow = []            

    # docname = row[0].split("/")[0]
    # arow.append(docname)
    for y in range(len(row)):
        listcoie = []
        if searchitem in row[y]:
            coie = ", ".join(row).split(divider)[1].split(end)[0]
            listcoie.append(coie)
            arow.append(listcoie)
            
        else:
            pass
    return arow



def generalSearch():
    term = raw_input("What term would you desire to search?: ")
    with open("copyallXLSX_DOCX.csv") as csvfile:
        readCSV = csv.reader(csvfile, delimiter=",")
        results =[]
        for row in readCSV:
            doc = row[0].split("/")[0]
            first = row[1]
            second = row[2]
            third = row[3]
            fourth = row[4]
         
            if term in first:
                print "Document ", doc, " has (1) ", first
                results.append(doc, first)
            elif term in second:
                print "Document ", doc, " has (2) ", second
                results.append(doc, second)
            elif term in third:
                print "Document ", doc, " has (3) ", third
                results.append(doc, third)
            elif term in fourth:
                print "Document ", doc, " has (4) ", fourth
                results.append(doc, fourth)
            else:
                print term, ' not found'
    return results

def createmodifedDates():
        #with open("copyallXLSX_DOCX.csv# ") as csvfile:
            # readCSV = csv.reader(csvfile, delimiter=",")
            # results =[]
            # for row in readCSV:
            #     if "created" or "modified" in row:
            #         print row[0], row[1]

        csv_file = csv.reader(open('copyallXLSX_DOCX.csv', "rb"), delimiter=",")
        rowlist=[]
        for row in csv_file:
            arow = []
            docname = row[0]
            arow.append(docname)
            if "dc:creator" in ", ".join(row):
                creator = ", ".join(row).split("dc:creator>")[1].split("<")[0]
            elif "dc:creator" not in ", ".join(row):
                creator = " - "
            arow.append(creator)
            if "lastModifiedBy>" in ", ".join(row):
                lmodby = ", ".join(row).split("lastModifiedBy>")[1].split("<")[0]
            else:
                lmodby = " - "
            arow.append(lmodby)
            if "dcterms:created" in ", ".join(row):
                crtm = ", ".join(row).split('dcterms:created xsi:type="dcterms:W3CDTF">')[1].split("<")[0]
            
            else:
                crtm = " - "
            arow.append(crtm)
            if "dcterms:modified " in ", ".join(row):
                modtm = ", ".join(row).split('dcterms:modified xsi:type="dcterms:W3CDTF">')[1].split("<")[0]
            else:
                modtm = " - "
            arow.append(modtm)
            if "Application>" in ", ".join(row):
                appl = ", ".join(row).split('Application>')[1].split("<")[0]
            else:
                appl = " - "
            arow.append(appl)
            if "DocSecurity>" in ", ".join(row):
                docsec = ", ".join(row).split('DocSecurity>')[1].split("<")[0]
            else:
                docsec = " - "
            arow.append(docsec)
            if "vt:lpstr>" in ", ".join(row):
                vtlp = row_multiple(row, 'vt:lpstr>', 'vt:lpstr>', "<")
            else:
                vtpl = " - "
            arow.append(vtpl)
            if "AppVersion>" in ", ".join(row):
                apv = ", ".join(row).split('AppVersion>')[1].split("<")[0]
            else:
                apv = " - "
            arow.append(apv)
            if "PK" in row[1]:
                pk = "Yes PK "
            else:
                pk = " - "
            arow.append(pk)
            if 'a:clrScheme name="' in ", ".join(row):
                #clrs = ", ".join(row).split('a:clrScheme name="')[1].split('"')[0]
                clrs = row_multiple(row, 'a:clrScheme name="','a:clrScheme name="', '"')
            else:
                clrs = " - "
            arow.append(clrs)
            if '/main" name="' in ", ".join(row):
                thnam = ", ".join(row).split('/main" name="')[1].split('"')[0]
            else:
                thnam = " - "
            arow.append(thnam)
            if 'ext uri="{' in ", ".join(row):
                #exturi = ", ".join(row).split('ext uri="{')[1].split('}')[0]
                exturi = row_multiple(row, 'ext uri="{', 'ext uri="{', '}')
            else:
                exturi = " - "
            arow.append(exturi)                        
            if 'http://schemas.microsoft.com/office/' in ", ".join(row):
                schmac = ", ".join(row).split('http://schemas.microsoft.com/office/')[1].split('/main"')[0]
            else:
                schmac = " - "
            arow.append(schmac)                        
            if 'perkinscoie.com</t' in ", ".join(row):
                listcoie = row_multiple(row, 'perkinscoie.com</t', '>', '</t')
                # listcoie = []
                # for y in len(row):
                #     if 'perkinscoie.com</t' in row[y]:
                #         coie = ", ".join(row).split('>')[1].split('</t')[0]
                #         listcoie.append(coie)
                #     else:
                #         pass
            else:
                listcoie = " - "
            arow.append(listcoie)                        
            if 'property fmtid="{' in ", ".join(row):
                prfmitd = ", ".join(row).split('property fmtid="{')[1].split('}')[0]
            else:
                prfmitd = " - "
            arow.append(prfmitd)                        
            if '}" pid="' in ", ".join(row):
                pid = ", ".join(row).split('}" pid="')[1].split('"')[0]
            else:
                pid = " - "
            arow.append(pid)
            if 'Company>' in ", ".join(row):
                company = ", ".join(row).split('Company>')[1].split('<')[0]
            else:
                company = " - "
            arow.append(company)
            if 'Template>' in ", ".join(row):
                template = ", ".join(row).split('Template>')[1].split('<')[0]
            else:
                template = " - "
            arow.append(template)
            if 'TotalTime>' in ", ".join(row):
                tottime = ", ".join(row).split('TotalTime>')[1].split('<')[0]
            else:
                tottime = " - "
            arow.append(tottime)
            if 'w:tplc="04090' in ", ".join(row):
                eng409 = "04090"
            else:
                eng409 = " - "
            arow.append(eng409)
 #####################           '<w:tmpl w:val="'
            if 'w:tmpl w:val="' in ", ".join(row):
                tmplcode = row_multiple(row, 'w:tmpl w:val="', 'w:tmpl w:val="', '"/')
            else:
                tmplcode = " - "
            arow.append(tmplcode)
            if 'w:nsid w:val="' in ", ".join(row):
                nsid = row_multiple(row, 'w:nsid w:val="', 'w:nsid w:val="', '/')
                #nsid = ", ".join(row).split('w:nsid w:val="')[1].split('"/')[0]
            else:
                nsid = " - "
            arow.append(nsid)
            if 'w:lang w:val="' in ", ".join(row):
                wlang = row_multiple(row, 'w:lang w:val="', 'w:lang w:val="', '"')
                #wlang = ", ".join(row).split('w:lang w:val="')[1].split('"')[0]
            else:
                wlang = " - "
            arow.append(wlang)
            if '" w:eastAsia="' in ", ".join(row):
                ealang = row_multiple(row, '" w:eastAsia="', '" w:eastAsia="', '"')
                #ealang = ", ".join(row).split('" w:eastAsia="')[1].split('"')[0]
            else:
                ealang = " - "
            arow.append(ealang)                        
            if '" w:bidi="' in ", ".join(row):
                bidilang = row_multiple(row, '" w:bidi="', '" w:bidi="', '"')
                #bidilang = ", ".join(row).split('" w:bidi="')[1].split('"')[0]
            else:
                bidilang = " - "
            arow.append(bidilang)
            if 'w:themeFontLang w:val="' in ", ".join(row):
                thfontlang = row_multiple(row, 'w:themeFontLang w:val="', 'w:themeFontLang w:val="', '"')
                #thfontlang = ", ".join(row).split('w:themeFontLang w:val="')[1].split('"')[0]
            else:
                thfontlang = " - "
            arow.append(thfontlang)
            if 'w:rsidRoot w:val="' in ", ".join(row):
                rsidroot = ", ".join(row).split('w:rsidRoot w:val="')[1].split('"')[0]
            else:
                rsidroot = " - "
            arow.append(rsidroot)
            if '<ds:datastoreItem ds:itemID="{' in ", ".join(row):
                datastoreid = ", ".join(row).split('<ds:datastoreItem ds:itemID="{')[1].split('}')[0]
            else:
                datastoreid = " - "
            arow.append(datastoreid)
            if 'StyleName="' in ", ".join(row):
                bib = row_multiple(row, 'StyleName="', 'StyleName="', '"')
                #bib = "APA Bibliography"
            else:
                bib = " - "
            arow.append(bib)
            if 'c:lang val="' in ", ".join(row):
                chartlang = ", ".join(row).split('c:lang val="')[1].split('"')[0]
            else:
                chartlang = " - "
            arow.append(chartlang)
            if 'dc:language>' in ", ".join(row):
                dclang = ", ".join(row).split('dc:language>')[1].split('<')[0]
            else:
                dclang = " - "
            arow.append(dclang)



            rowlist.append(arow)
           
        return rowlist



def main_write():
    for row in createmodifedDates():
        with open('DOCXxlsxCreateModTimes.csv', 'ab') as csvfile:
            spamwriter = csv.writer(csvfile, dialect='excel')
            spamwriter.writerow(row)


def getlist():
    with open('DOCXxlsxCreateModTimes.csv', 'rb') as csvfile:
        rcsv = csv.reader(csvfile, delimiter=",")
        namelist = []
        for row in rcsv:
            filename = row[0].split("/")[0]
            if filename in ", ".join(namelist):
                pass
            else:
                namelist.append(filename)
        return namelist

#given a document name loops through the csv and is name is in line of csv adds line to list and returns the list
def myloop(nl):
    with open('DOCXxlsxCreateModTimes.csv', 'rb') as csvfile:
        rcsv = csv.reader(csvfile, delimiter=",")
        singledoclist = []
        singledoclist.append(nl)
        for row in rcsv:
            if nl in row[0].split("/")[0]:
                singledoclist.append(row)
            else:
                pass
        return singledoclist

#takes document name list from getlist() and loops through it, calling myloop with each document name, which returns a list of each document data. This is added to docxdata list.  
def csv_tidy():
    
        docxdata = []
        namelist = getlist()
        #namelist= ["2-19-16-friends-of-hrc-list_hfa16-giving-history", "4-16-commitment-sheet_040416-update", "7-1-15-commitment-sheet", "donors"]

        for x in range(0, len(namelist)):
            docxdata.append(myloop(namelist[x]))
        print "Done"
        print ""
        with open('FINAL_DOCX_XLSX_DATA.csv', 'ab') as csvfile:
            spamwriter = csv.writer(csvfile, dialect='excel')
            
            for line in docxdata:
                myline = []
                myline.append(line[0])
                for x in line[1:]:
                    for y in x:
                        myline.append(y)
            
                spamwriter.writerow(myline) 
            print " RESULTS WRITTEN TO: FINAL_DOCX_XLSX_DATA.csv"


        
###########################################################
# MAIN LOOP                                               #
# 1 WRITE THE DOCUMENT WITH main_write,                   #
# 2 then tidy by csv_tidy after supplying a list of files #
###########################################################


#DONE: main_write() 
#DONE getlist()
csv_tidy()
