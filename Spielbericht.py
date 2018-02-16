import PyPDF2
import openpyxl
import operator
import pickle

version = '''Version 0.4'''

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
        i = 2
        numTrainer = 0
        rawText = rawText[pos_Manschaft + len(name) + 78:]

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
                spieler_name += rawText[i]
                self.trainer.append(spieler(spieler_name, chr(ord('A')+numTrainer)))
                numTrainer += 1
            elif (rawText[i].isalpha() and not rawText[i + 1].isupper()):
                debug = rawText[i-2:]
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
        for i in range(0, len(self.players)-1):
            if self.players[i].number > self.players[i+1].number:
                for u in range(0, i+1):
                    self.players[u].number = self.players[u].number % 10
        for each in self.players[:]:
            if each.number in [1, 12, 16]:
                self.players.remove(each)
                self.torwart.append(each)
                self.torwart.sort(key=operator.attrgetter('number'))
    def changeNumberPlayer(self, player, number):
        if not (player in self.players or self.torwart): return False
        oldNumber = player.number
        player.setNumber(number)
        if number not in [1, 12, 16]:
            self.players.sort(key=operator.attrgetter('number'))
        else:
            if oldNumber in [1, 12, 16]:
                self.torwart.remove(player)
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
        akt = input('Wollen sie die Manschaft ' + str(name)+ ' aktualiesieren?(j/n)')
        if akt == 'j':
            del Manschaften[name]
            Manschaften.update({name:manschaft(name, file)})
    else:
        Manschaften.update({name: manschaft(name, file)})
        Manschaften_kurz.update({input('Kürzel für Manschaft: '+ str( name)): name})

    pos_Gast = rawText.find('Gast: ')
    pos_Ende = rawText.find('Nr.Name', pos_Gast)
    name = rawText[pos_Gast + 6:pos_Ende]
    if name in Manschaften.keys():
        akt = input('Wollen sie die Manschaft '+ str(name) + ' aktualiesieren?(j/n)')
        if akt == 'j':
            del Manschaften[name]
            Manschaften.update({name: manschaft(name, file)})
    else:
        Manschaften.update({name: manschaft(name, file)})
        Manschaften_kurz.update({input('Kürzel für Manschaft: ' + str(name)): name})


def fileSchreiben():
    global Manschaften, Manschaften_kurz
    heim = input('Heimmanschaft?')
    if heim == '':
        heim = 'HSG Böblingen/Sindelfingen'
        print('Heim automatisch auf', heim, 'gesetzt')
    if heim not in Manschaften.keys():
        if heim not in Manschaften_kurz.keys():
            print('Manschaft nicht bekannt!')
            return
        else:
            heim = Manschaften_kurz[heim]
    gast = input('Gastmanschaft?')
    if gast not in Manschaften.keys():
        if gast not in Manschaften_kurz.keys():
            print('Manschaft nicht bekannt!')
            return
        else:
            gast = Manschaften_kurz[gast]
    sheet['A2'] = 'Aufstellung vom ' + input('Datum?')
    heimManschaft = Manschaften[heim]
    gastManschaft = Manschaften[gast]
    wb = openpyxl.load_workbook('Mannschaftsliste_MUSTER.xlsx')
    sheet = wb[wb.sheetnames[0]]
    sheet['A1'] = heimManschaft.spielklasse
    sheet['A3'] = heimManschaft.name
    sheet['F3'] = gastManschaft.name
    for i in range(0, len(heimManschaft.torwart)):
        sheet['B' + str(7+i)] = heimManschaft.torwart[i].name
        sheet['A' + str(7+i)] = heimManschaft.torwart[i].number
        sheet['D' + str(7+1)] = 'TW'
    for i in range(0, len(heimManschaft.players)):
        sheet['B' + str(10 + i)] = heimManschaft.players[i].name
        sheet['A' + str(10+i)] = heimManschaft.players[i].number
    for i in range(0, len(heimManschaft.trainer)):
        sheet['B' + str(27 + i)] = heimManschaft.trainer[i].name
    for i in range(0, len(gastManschaft.torwart)):
        sheet['G' + str(7+i)] = gastManschaft.torwart[i].name
        sheet['F' + str(7+i)] = gastManschaft.torwart[i].number
        sheet['I' + str(7+1)] = 'TW'
    for i in range(0, len(gastManschaft.players)):
        sheet['G' + str(10 + i)] = gastManschaft.players[i].name
        sheet['F' + str(10+i)] = gastManschaft.players[i].number
    for i in range(0, len(gastManschaft.trainer)):
        sheet['G' + str(27 + i)] = gastManschaft.trainer[i].name
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

print('Help(h)')
print('Quit(q)')
print('Bogen kreieren(b)')
print('Datei lesen(d)')
print('Kürzel liste(k)')
print('Nummer ändern(n)')
print('Manschaftsliste(m)')
while True:
    cmd = input('->')
    if cmd == 'q':
        print('Godbye!')
        break
    elif cmd == 'h':
        print('Help(h)')
        print('Quit(q)')
        print('Bogen kreieren(b)')
        print('Datei lesen(d)')
        print('Kürzel liste(k)')
        print('Nummer ändern(n)')
        print('Manschaftsliste(m)')
    elif cmd == 'm':
        akt_manschaft = input('Von welcher Manschaft wollen Sie die Manschaftsliste?')
        if akt_manschaft not in Manschaften.keys():
            if akt_manschaft not in Manschaften_kurz.keys():
                print('Manschaft nicht bekannt!')
                continue
            else:
                akt_manschaft = Manschaften_kurz[akt_manschaft]
        akt_manschaft = Manschaften[akt_manschaft]
        for each in akt_manschaft.torwart:
            print('{:>3}|{:20}|TW'.format(each.number, each.name))
        for each in akt_manschaft.players:
            print('{:>3}|{:30}|'.format(each.number, each.name))
        print('-'*35)
        for each in akt_manschaft.trainer:
            print('{:>3}|{:30}|'.format(each.number, each.name))
    elif cmd == 'n':
        akt_manschaft = input('Von welcher Manschaft ist der Spieler bei dem sie die Nummer ändern wollen?')
        if akt_manschaft not in Manschaften.keys():
            if akt_manschaft not in Manschaften_kurz.keys():
                print('Manschaft nicht bekannt!')
                continue
            else:
                akt_manschaft = Manschaften_kurz[akt_manschaft]
        akt_manschaft = Manschaften[akt_manschaft]
        name = input('Wie heißt der Spieler?')
        new_number = int(input('Welche nummer soll der Spieler bekommen?'))
        tempbool = False
        for i in range(0, len(akt_manschaft.players)):
            if akt_manschaft.players[i].name == name:
                akt_manschaft.changeNumberPlayer(akt_manschaft.players[i], new_number)
                tempbool = True
        for i in range(0, len(akt_manschaft.torwart)):
            if akt_manschaft.torwart[i].name == name:
                akt_manschaft.changeNumberPlayer(akt_manschaft.torwart[i], new_number)
                tempbool = True
        if tempbool:
            print('Erfolgreich Nummer von', name, 'auf', new_number, 'gesetzt.')
        else:
            print('Spieler nicht gefunden!')
    elif cmd == 'k':
        for each in Manschaften_kurz.keys():
            print(each)
            print(Manschaften_kurz[each])
            print(Manschaften[Manschaften_kurz[each]].spielklasse, end='\n\n')
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