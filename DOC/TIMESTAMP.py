from filetimes import *



#in = raw hex. out = list split at d101 = 2015/16
# note: very basic - have to filter after this
def timestamp(hexdata):
    tstamplistd101=[]
    chunks=hexdata.split("d101")
    for ts in chunks:
        tstamplistd101.append(ts[-12:]+"d101")
    dunks=hexdata.split("d102")
    for st in dunks:
        tsstamplistd101.append(st[-12:]+"d102")
    return tstamplistd101


#In = list of 16 char timestamps in BE order. Out= call Win32Time function supplied with list in LE order 
def sortTimeStamps(tslist):
    for ts in tslist:
        infoMeta.append(ts)
        hightime = ts[6:8]+ts[4:6]+ts[2:4]+ts[0:2]
        lowtime =ts[14:16]+ts[12:14]+ts[10:12]+ts[8:10]
        ft= hightime+":"+lowtime
        infoMeta.append(ft)
        return ft



def win32time(ft):
    # switch parts
    h2, h1 = [int(h, base=16) for h in ft.split(':')]
    # rebuild
    ft_dec = struct.unpack('>Q', struct.pack('>LL', h1, h2))[0]
    # add ft_dec to infometa
    infoMeta.append(ft_dec)
    # use function from iceaway's comment
    infoMeta.append(filetime_to_dt(ft_dec))


# to do get the main stuff a workin'



def Main(url):
    print "Really about to dl: ", url
    file_name, hexdata = Downloadfile(url)
    tstamplistd101 = timestamp(hexdata)
    ft = sortTimeStamps(tstamplistd101)
    win32time(ft)
    aline=[]
    for item in infoMeta:
        aline.append(item)
    with open("mydocs.csv", "w") as f:
        an_entry = csv.writer(f)
        an_entry.writerow(aline)
    f.close()


#url = "https://guccifer2.files.wordpress.com/2016/07/hsu-contributions.xls"
#Main(url)

with open("attachments.csv", "r") as atts:
    attline = csv.reader(atts)

    for line in attline:
        if "doc" in line[1] and "docx" not in line[1]:
            print "Downloading: ", line[1]
            Main(line[1])
        else:
            pass
    atts.close()


    

