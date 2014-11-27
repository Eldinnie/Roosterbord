'''
Created on 21 nov. 2014

@author: Pieter
'''
from bs4 import BeautifulSoup, NavigableString
import urllib2, datetime, sys
import xlsxwriter
import re
from lesroosters import lesroosters

maanden={"januari":1,"februari":2,"maart":3,"april":4,"mei":5,"juni":6,"juli":7,
         "augustus":8,"september":9,"oktober":10,"november":11,"december":12}
wijzmatch = re.compile(r"^([0-9]{1}[A-Z]{1}) ([A-Z]+) \([A-Z]{1}[a-z]{1}\)\*? (.{3})$")
wijzmatch2 = re.compile(r"^([0-9]+)([A-Z]+) \([A-Z]{1}[a-z]{1}\) (.{3})$")


class lesdag():
    def __init__(self,datum, bijgewerkt, afwezigen, wijzigingen, vrij, med):
        self.rooster = {1:{},2:{},3:{},4:{},5:{},6:{},7:{},8:{}}
        self.mededelingen=[]
        if med[0] != u"":
            self.mededelingen.append("Hele dag "+med[0])
        self.datum = self._make_date(datum)
        self.bijgewerkt = self._make_date(bijgewerkt)
        self.afwezigen = afwezigen
        self._compile_vrij(vrij)
        self._compile_wijzingen(wijzigingen)


    def _make_date(self,datum):
        tmp = re.match("^Bijgewerkt op ([0-9]+) ([a-zA-Z]+) ([0-9]{4}) om ([0-9]{2}):([0-9]{2}) uur$", datum.strip())
        if tmp:
            tmp = datetime.datetime(int(tmp.groups()[2]),maanden[tmp.groups()[1]],int(tmp.groups()[0]), int(tmp.groups()[3]),int(tmp.groups()[4]))
        else:
            tmp = re.match("^[a-zA-Z]+ ([0-9]+) ([a-zA-Z]+) ([0-9]{4})$", datum.strip())
            tmp= datetime.datetime(int(tmp.groups()[2]),maanden[tmp.groups()[1]],int(tmp.groups()[0]))
        return tmp

    def getDateString(self, hours=False):
        if hours:
            return self.bijgewerkt.strftime("%A %d %B - %H:%M")
        else:
            return self.datum.strftime("%A %d %B")


    def _compile_vrij(self, vrij):
        for uur,vrij in enumerate(vrij):
            self.rooster[uur+1]["vrij"] = vrij

    def _compile_wijzingen(self,wijz):
        for uur,wij in enumerate(wijz):
            tmp= []
            for wi in wij.contents:
                if type(wi) == NavigableString and wi.strip() != u"":
                    try:
                        if wijzmatch.match(wi.strip()):
                            klas, vak, lokaal = wijzmatch.match(wi.strip()).groups()
                        else:
                            klas, vak, lokaal = wijzmatch2.match(wi.strip()).groups()
                        tmp.append({'klas':klas,'vak':vak,'lokaal':lokaal})
                    except AttributeError:
                        self.mededelingen.append("Uur "+str(uur+1)+" "+wi.strip())
            self.rooster[uur+1]["wijzigingen"] = tmp

def printRooster(ros):
    print ros['naam']+"\n"+"-"*20
    print "Uur\tMa\tDi\tWo\tDo\tVrij"
    for x in range(1,9):
        tmp=[]
        tmp.append(str(x))
        for y in range(5):
            tmp.append("%s" % ((ros[y][x]['klas']+" "+ros[y][x]['lokaal'])))
        print "".join(["|%s\t" % (q) for q in tmp])

def UpdateFromWeb():
    raw = urllib2.urlopen("http://www.gymnasiumhilversum.nl/rooster/dagrooster.html").read()
    raw,_= re.subn('</BR>', "<BR />", raw)
    return raw

def UpdateFromFile():
    with open("rooster.txt","r") as f:
        raw = f.read()
    return raw

def processRooster(raw):
    soup = BeautifulSoup(raw,"lxml")
#     with open("rooster.txt","w") as f:
#         f.write("Updatet "+str(datetime.datetime.now())+"\n")
#         f.write(soup.prettify())
    dates = [x.string for x in soup.find_all("td",{"class":"datum"})]
    bijgewerkt = [x.string for x in soup.find_all("td",{"class":"bijgewerkt"})]
    afwezigen = [x.string.strip() for x in soup.find_all("td",{"class":"afwezig"})]
    lessen = [x for x in soup.find_all("td",{"class":"les"})]
    vrij = [x.string.strip().split(", ") for x in soup.find_all("td",{"class":"vrij"})]
    med = [x.string.strip().split(", ") for x in soup.find_all("td",{"class":"mededeling"})]
    lesarr = []
    lesarr.append(lessen[:8:2]+lessen[1:8:2])
    lesarr.append(lessen[8::2]+lessen[9::2])
    vrijarr=[]
    vrijarr.append(vrij[:8:2]+vrij[1:8:2])
    vrijarr.append(vrij[8::2]+vrij[9::2])
    dag1 = lesdag(dates[0],bijgewerkt[0],afwezigen[0], lesarr[0], vrijarr[0], med[0])
    dag2 = lesdag(dates[1],bijgewerkt[1],afwezigen[1], lesarr[1], vrijarr[1], med[1])
    return dag1, dag2

def checkDayToRooster(dag,rooster):
    roosterdag = rooster[dag.datum.weekday()]
    meldingen = ""
    for uur,data in dag.rooster.items():
#         print uur, data, roosterdag[uur]
        try:
            if roosterdag[uur]['klas'] in data['vrij']:
                meldingen += "%s Uur %d is klas %s vrij ipv les in %s\n" % (dag.getDateString(), uur, roosterdag[uur]['klas'], roosterdag[uur]['lokaal'])
                rooster[dag.datum.weekday()][uur]['klas']= "!"+rooster[dag.datum.weekday()][uur]['klas']
        except KeyError:
            pass
        for wijz in data['wijzigingen']:
            try:
                if wijz['klas'] == roosterdag[uur]['klas']:
                    meldingen += "%s Uur %d heeft klas %s %s in lokaal %s ipv van NA in %s\n" % (dag.getDateString(), uur, wijz['klas'], wijz['vak'], wijz['lokaal'], roosterdag[uur]['lokaal'])
                    rooster[dag.datum.weekday()][uur]['klas']= "!"+rooster[dag.datum.weekday()][uur]['klas']
            except KeyError:
                pass
            try:
                if wijz['vak'] == roosterdag[uur]['vak'] and wijz['klas'] != roosterdag[uur]['klas']:
                    meldingen += "%s uur %d heeft klas %s %s in lokaal %s. Rooster zegt %s %s in lokaal %s\n" % (dag.getDateString(), uur, wijz['klas'], wijz['vak'],  wijz['lokaal'], roosterdag[uur]['klas'], roosterdag[uur]['vak'], roosterdag[uur]['lokaal'])
                    rooster[dag.datum.weekday()][uur]['klas']= "!"+rooster[dag.datum.weekday()][uur]['klas']
            except KeyError:
                pass
    meldingen += "\n".join(["%s %s" % (dag.getDateString(), x) for x in dag.mededelingen])
    meldingen +="\n"
    return rooster, meldingen



# Create an new Excel file and add a worksheet.

def makeExcell(bestand, data, bijgewerkt):
    workbook = xlsxwriter.Workbook(bestand)
    for rooster,meldingen in data:
        worksheet = workbook.add_worksheet(rooster['naam'])
        worksheet.set_column('A:F', 15)
        bold = workbook.add_format({'bold': True})
        red_shade = workbook.add_format({'bg_color': '#FFC7CE',
                                   'font_color': '#9C0006'})

        worksheet.write('A1', rooster['naam'])
        worksheet.write('B1', bijgewerkt)
        worksheet.write('A2', 'Uur', bold)
        worksheet.write('B2', 'Maandag', bold)
        worksheet.write('C2', 'Dinsdag', bold)
        worksheet.write('D2', 'Woensdag', bold)
        worksheet.write('E2', 'Donderdag', bold)
        worksheet.write('F2', 'Vrijdag', bold)

        start_row = 2
        start_col = 0

        for x in range(0,8):
            worksheet.write(x+start_row, 0, str(x+1))
            for y in range(1,6):
                if rooster[y-1][x+1]['klas'][:1]=="!":
                    worksheet.write(x+start_row, y+start_col, rooster[y-1][x+1]['klas'][1:]+" "+rooster[y-1][x+1]['vak']+" "+rooster[y-1][x+1]['lokaal'],red_shade)
                else:
                    worksheet.write(x+start_row, y+start_col, rooster[y-1][x+1]['klas']+" "+rooster[y-1][x+1]['vak']+" "+rooster[y-1][x+1]['lokaal'])

        for x, melding in enumerate(meldingen.split("\n")):
            worksheet.write(x+start_row+9, 0, melding)
    workbook.close()

if __name__=="__main__":
    if len(sys.argv)>1:
        bestand = sys.argv[1]
    else:
        bestand = "rooster.xlsx"
    raw = UpdateFromWeb()
    dag1, dag2 = processRooster(raw)
    voorExcell = []
    for rooster in lesroosters:
        rooster, meldingen =  checkDayToRooster(dag1, rooster)
        rooster, meldingen2 = checkDayToRooster(dag2, rooster)
        meldingen += meldingen2
#         printRooster(rooster)
#         print "bijgwerkt op: "+str(dag1.bijgewerkt)
#         print meldingen or "Niets te melden"
        voorExcell.append((rooster, meldingen))
        bijgewerkt = dag1.getDateString(hours=True)
    makeExcell(bestand, voorExcell, bijgewerkt)