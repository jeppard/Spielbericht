import PyPDF2
import openpyxl
import operator
import pickle

version = '''Version 0.2'''

class spieler():
    def __init__(self, name, number=None):
        self.name = name
        self.number = number
    def setNumber(self, number):
        self.number = number

class manschaft():
    def __init__(self, name, file):
        self.name = name
        pdfFile = open(file, 'rb')
        pdfReader = PyPDF2.PdfFileReader(pdfFile)
        pageObj = pdfReader.getPage(1)
        pageOne = pdfReader.getPage(0)
        p1 = pageOne.extractText()
        start = p1.find('Spielklasse') + 11
        ende = p1.find('Spiel/Datum')
        self.spielklasse = p1[start:ende]
        rawText = pageObj.extractText()
        pos_Manschaft = rawText.find(name)
        self.players = []
        self.trainer = []
        self.torwart = []
        i = 0
        numTrainer = 0
        rawText = rawText[pos_Manschaft + len(name) + 80:]

        while True:
            if rawText[i:i + 4] == 'Gast' or rawText[i:i + 8] == 'Handball':
                break
            if (rawText[i] in ['A', 'B', 'C', 'D'] and
                    rawText[i + 1].isupper() and rawText[i + 2].islower() and
                    not (rawText[i:i + 6] == 'DGast:' or rawText[i:i + 9] == 'DHandball')):
                i += 1
                spieler_name = ''
                while (rawText[i].isalpha() and not rawText[i + 1].isupper()) or rawText[i] == ' ':
                    spieler_name += rawText[i]
                    i += 1
                self.trainer.append(spieler(spieler_name, chr(ord('A')+numTrainer)))
                numTrainer += 1
            elif (rawText[i].isalpha() and not rawText[i + 1].isupper()):
                if rawText[i-2].isnumeric():
                    number = int(rawText[i-2:i])
                else:
                    number = int(rawText[i-1])
                spieler_name = ''
                while (rawText[i].isalpha() and not rawText[i + 1].isupper()) or rawText[i] == ' ':
                    spieler_name += rawText[i]
                    i += 1
                self.players.append(spieler(spieler_name, number))
            i += 1
            for i in range(0, len(self.players)):
                if self.players[i].number > self.players[i+1].number:
                    for u in range(0, i):
                        self.players[u].number = self.players[u].number % 10
            for each in self.players[:]:
                if each.number in [1, 12, 16]:
                    self.players.remove(each)
                    self.torwart.append(each)
                    self.torwart.sort(key=operator.attrgetter('number'))
    def changeNumberPlayer(self, player, number):
        if not (player in self.players or self.torwart): return False
        player.setNumber(number)
        if number not in [1, 12, 16]:
            self.players.sort(key=operator.attrgetter('number'))
        else:
            self.players.remove(player)
            if not (player in self.torwart):
                self.torwart.append(player)
            self.torwart.sort(key=operator.attrgetter('number'))
        return True

def fileRead(file):
    global Manschaften_kurz, Manschaften
    pdfFile = open(file, 'rb')
    pdfReader = PyPDF2.PdfFileReader(pdfFile)
    pageObj = pdfReader.getPage(1)
    rawText = pageObj.extractText()
    pos_Heim = rawText.find('Heim: ')
    pos_Ende = rawText.find('Nr.Name')
    name = rawText[pos_Heim + 6:pos_Ende]
    if name in Manschaften.keys():
        akt = input('Wollen sie die Manschaft', name, 'aktualiesieren?(j/n)')
        if akt == 'j':
            del Manschaften[name]
            Manschaften.update({name:manschaft(name, file)})
    else:
        Manschaften.update({name: manschaft(name, file)})
        Manschaften_kurz.update({input('Kürzel für Manschaft:', name): name})

def fileSchreiben():
    global Manschaften, Manschaften_kurz
    heim = input('Heimmanschaft?')
    if heim == '':
        heim = 'HSG Böblingen/Sindelfingen'
        print('Heim automatisch auf', heim, 'gesetzt')
    if heim not in Manschaften.keys():
        if heim not in Manschaften_kurz.keys():
            print('Manschaft nicht bekannt!')
        else:
            heim = Manschaften_kurz[heim]
    gast = input('Gastmanschaft?')
    if gast not in Manschaften.keys():
        if gast not in Manschaften_kurz.keys():
            print('Manschaft nicht bekannt!')
        else:
            gast = Manschaften_kurz[gast]
    heimManschaft = Manschaften[heim]
    gastManschaft = Manschaften[gast]
    wb = openpyxl.load_workbook('Mannschaftsliste_MUSTER.xlsx')
    sheet = wb[wb.sheetnames[0]]
    sheet['A3'] = heimManschaft.name
    sheet['F3'] = gastManschaft.name
    for i in range(0, len(heimManschaft.players)):
        sheet['B' + str(7 + i)] = heimManschaft.players[i]
    for i in range(0, len(heimManschaft.trainer)):
        sheet['B' + str(27 + i)] = heimManschaft.trainer[i]
    for i in range(0, len(gastManschaft.players)):
        sheet['G' + str(7 + i)] = gastManschaft.players[i]
    for i in range(0, len(gastManschaft.trainer)):
        sheet['G' + str(27 + i)] = gastManschaft.trainer[i]
    datei = input('Dateiname:')
    try:
        if datei[-5:0] != '.xlsx':
            datei += '.xlsx'
    except:
        datei += '.xlsx'
    wb.save(datei)


print(version)
try:
    Manschaften = pickle.load(open('save.p', 'rb'))
    Manschaften_kurz = pickle.load(open('save2.p', 'rb'))
except:
    print('Manschaften konnten nicht geladen werden!')
    Manschaften = {}
    Manschaften_kurz = {}

while True:
    cmd = input('->')
    if cmd == 'q':
        break
    elif cmd == 'h':
        print('Help(h)')
        print('Quit(q)')
        print('Bogen kreieren(b)')
        print('Datei lesen(d)')
        print('Kürzel liste(k)')
    elif cmd == 'k':
        for each in Manschaften_kurz.keys():
            print(each)
            print(Manschaften_kurz[each], end='\n\n')
    elif cmd == 'b':
        fileSchreiben()
    elif cmd == 'd':
        datei = input('Dateiname?')
        if datei[-4:] != '.pdf':
            datei += '.pdf'
        fileRead(datei)
    else:
        print('Befehl unbekannt!\nh für Hilfe')
    
pickle.dump(Manschaften, open('save.p', 'wb'))
pickle.dump(Manschaften_kurz, open('save2.p', 'wb'))