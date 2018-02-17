import PyPDF2
import openpyxl
import operator
import pickle

version = '''Version 0.9'''

class spieler():
    def __init__(self, name, number=None):
        self.name = name
        self.number = number
    def setNumber(self, number):
        self.number = number

class manschaft():
    def __init__(self):
        self.name = None
        self.spielklasse = None
        self.players = []
        self.trainer = []
        self.torwart = []
    def read(self, name, file):
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
                    if not rawText[i-1].isnumeric():
                        print('Manschaft Fehlerhaft')
                        number = 100
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
            elif each.number == 100:
                self.players.remove(each)
                self.trainer.append(spieler(each.name[1:], each.name[0]))
    def changeNumberPlayer(self, player, number):
        if not (player in self.players or self.torwart or self.trainer): return False
        oldNumber = player.number
        player.setNumber(number)
        if number not in [1, 12, 16] or not number.isalpha:
            if oldNumber in [1, 12, 16]:
                self.torwart.remove(player)
                self.players.append(player)
            elif oldNumber.isaplha():
                self.trainer.remove(player)
                self.players.append(player)
            self.players.sort(key=operator.attrgetter('number'))
        elif number.isalpha():
            if oldNumber in [1, 12, 16]:
                self.torwart.remove(player)
                self.trainer.append(player)
            elif oldNumber.isnumeric():
                self.players.remove(player)
                self.trainer.append(player)
            self.trainer.sort(key=operator.attrgetter('number'))
        else:
            if oldNumber.isaplha:
                self.trainer.remove(player)
                self.torwart.append(player)
            elif oldNumber.isnumeric() and oldNumber not in [1, 12, 16]:
                self.players.remove(player)
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
    pageOne = pdfReader.getPage(0)
    p1 = pageOne.extractText()
    start = p1.find('Spielklasse') + 11
    ende = p1.find('Spiel/Datum')
    spielklasse = p1[start:ende]
    if name in Manschaften.keys():
        if spielklasse == Manschaften[name].spielklasse:
            akt = input('Wollen sie die Manschaft ' + str(name)+ ' aktualisieren?(j/n)')
            if akt == 'j':
                del Manschaften[name]
                Manschaften.update({name:manschaft()})
                Manschaften[name].read(name, file)
        else:
            Manschaften.update({name: manschaft()})
            Manschaften[name].read(name, file)
            Manschaften_kurz.update({input('Kürzel für Manschaft: ' + str(name)): name})
    else:
        Manschaften.update({name: manschaft()})
        Manschaften[name].read(name, file)
        Manschaften_kurz.update({input('Kürzel für Manschaft: '+ str( name)): name})

    pos_Gast = rawText.find('Gast: ')
    pos_Ende = rawText.find('Nr.Name', pos_Gast)
    name = rawText[pos_Gast + 6:pos_Ende]
    if name in Manschaften.keys():
        akt = input('Wollen sie die Manschaft '+ str(name) + ' aktualisieren?(j/n)')
        if akt == 'j':
            del Manschaften[name]
            Manschaften.update({name: manschaft()})
            Manschaften[name].read(name, file)
    else:
        while True:
            kuerzel = input('Kürzel für Manschaft: ' + str(name))
            if kuerzel in Manschaften_kurz:
                print('Kürzel schon vorhanden!')
            else:
                Manschaften_kurz.update({kuerzel: name})
                Manschaften.update({name: manschaft()})
                Manschaften[name].read(name, file)
                break


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
    heimManschaft = Manschaften[heim]
    gastManschaft = Manschaften[gast]
    wb = openpyxl.load_workbook('Mannschaftsliste_MUSTER.xlsx')
    sheet = wb[wb.sheetnames[0]]
    sheet['A1'] = heimManschaft.spielklasse
    sheet['A3'] = heimManschaft.name
    sheet['F3'] = gastManschaft.name
    sheet['A2'] = 'Aufstellung vom ' + input('Datum?')
    for i in range(0, len(heimManschaft.torwart)):
        sheet['B' + str(7+i)] = heimManschaft.torwart[i].name
        sheet['A' + str(7+i)] = heimManschaft.torwart[i].number
        sheet['D' + str(7+i)] = 'TW'
    for i in range(0, len(heimManschaft.players)):
        sheet['B' + str(10 + i)] = heimManschaft.players[i].name
        sheet['A' + str(10+i)] = heimManschaft.players[i].number
    for i in range(0, len(heimManschaft.trainer)):
        sheet['B' + str(27 + i)] = heimManschaft.trainer[i].name
    for i in range(0, len(gastManschaft.torwart)):
        sheet['G' + str(7+i)] = gastManschaft.torwart[i].name
        sheet['F' + str(7+i)] = gastManschaft.torwart[i].number
        sheet['I' + str(7+i)] = 'TW'
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
print('Manschaftsliste(m)')
print('Speichern(s)')
print('Manschaften editieren(e)')
print('Manschaften löschen(l)')
print('Manschaft hinzufügen(h)    (Manuell)')
print('Kürzel ändern(ä)')
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
        print('Manschaftsliste(m)')
        print('Speicher(s)')
        print('Manschaft editieren(e)')
        print('Manschaften löschen(l)')
        print('Manschaft hinzufügen(h)    (Manuell)')
        print('Kürzel ändern(ä)')
    elif cmd == 'ä':
        akt_manschaft = input('Von welcher Manschaft soll das kürzel geändert werden?')
        kuerzel = input('Altes Kürzel der Manschaft?')
        if akt_manschaft not in Manschaften.keys():
            if akt_manschaft not in Manschaften_kurz.keys():
                print('Manschaft nicht bekannt!')
                continue
            else:
                akt_manschaft = Manschaften_kurz[akt_manschaft]
        if akt_manschaft != Manschaften_kurz[kuerzel]:
            print('Kürzel Falsch!')
            continue
        new_kuerzel = input('Neues Kürzel?')
        if new_kuerzel in Manschaften_kurz.keys():
            print('Kürzel schon vorhanden!')
        else:
            del Manschaften_kurz[name]
            Manschaften_kurz.update({new_kuerzel:akt_manschaft})
            print('Erfolgreich Kürzel der Manschschaft', akt_manschaft, 'von', kuerzel, 'auf', new_kuerzel)
    elif cmd == 'h':
        name = input('Manschaftname?')
        kuerzel = input('Kürzel?')
        if kuerzel in Manschaften_kurz.keys():
            print('Kürzel schon vergeben!')
            continue
        spielklasse = input('Spielklasse?')
        print('Wollen sie wirklich die Manschaft', name, 'mit dem Kürzel', kuerzel, 'in der Spielklasse', spielklasse, 'kreieren',end='')
        temp =input('(j/n)')
        if temp =='j':
            Manschaften_kurz.update({kuerzel:name})
            Manschaften.update({name:manschaft()})
            Manschaften[name].name = name
            Manschaften[name].spielklasse = spielklasse
    elif cmd == 'l':
        akt_manschaft = input('Welche Manschaft soll gelöscht werden?')
        kuerzel = input('Kürzel der Manschaft?')
        if akt_manschaft not in Manschaften.keys():
            if akt_manschaft not in Manschaften_kurz.keys():
                print('Manschaft nicht bekannt!')
                continue
            else:
                akt_manschaft = Manschaften_kurz[akt_manschaft]
        if akt_manschaft != Manschaften_kurz[kuerzel]:
            print('Kürzel Falsch!')
            continue
        temp = input('Wollen sie wirklich die Manschaft '+str(akt_manschaft.name)+' löschen(j/n)')
        if temp == 'j':
            del Manschaften_kurz[kuerzel]
            del Manschaften[akt_manschaft]
    elif cmd == 'e':
        akt_manschaft = input('Welche Manschaft soll editiert werden?')
        if akt_manschaft not in Manschaften.keys():
            if akt_manschaft not in Manschaften_kurz.keys():
                print('Manschaft nicht bekannt!')
                continue
            else:
                akt_manschaft = Manschaften_kurz[akt_manschaft]
        akt_manschaft = Manschaften[akt_manschaft]
        print('Was willst du machen?')
        auswahl = input('Spieler Löschen(l)\nSpilernummer ändern(n)\nSpieler hinzufügen(h)\nSpielername ändern(ä)')
        if auswahl == 'n':
            name = input('Wie heißt der Spieler?')
            new_number = input('Welche nummer soll der Spieler bekommen?')
            tempbool = False
            for i in range(0, len(akt_manschaft.players)):
                if akt_manschaft.players[i].name == name:
                    akt_manschaft.changeNumberPlayer(akt_manschaft.players[i], int(new_number))
                    tempbool = True
            for i in range(0, len(akt_manschaft.torwart)):
                if akt_manschaft.torwart[i].name == name:
                    akt_manschaft.changeNumberPlayer(akt_manschaft.torwart[i], int(new_number))
                    tempbool = True
            for i in range(0, len(akt_manschaft.trainer)):
                if akt_manschaft.torwart[i].name == name:
                    akt_manschaft.changeNumberPlayer(akt_manschaft.trainer[i], new_number)
                    tempbool = True
            if tempbool:
                print('Erfolgreich Nummer von', name, 'auf', new_number, 'gesetzt.')
            else:
                print('Spieler nicht gefunden!')
        elif auswahl == 'ä':
            name = input('Wie heißt der Spieler?')
            new_name = input('Welche Namen soll der Spieler bekommen?')
            tempbool = False
            for i in range(0, len(akt_manschaft.players)):
                if akt_manschaft.players[i].name == name:
                    akt_manschaft.players[i].name = new_name
                    tempbool = True
            for i in range(0, len(akt_manschaft.torwart)):
                if akt_manschaft.torwart[i].name == name:
                    akt_manschaft.torwart[i].name = new_name
                    tempbool = True
            for i in range(0, len(akt_manschaft.trainer)):
                if akt_manschaft.trainer[i].name == name:
                    akt_manschaft.trainer[i].name = new_name
                    tempbool = True
            if tempbool:
                print('Erfolgreich Nummer von', name, 'auf', new_number, 'gesetzt.')
            else:
                print('Spieler nicht gefunden!')
        elif auswahl == 'l':
            name = input('Wie heißt der Spieler?')
            tempbool = False
            for i in range(0, len(akt_manschaft.players)):
                if akt_manschaft.players[i].name == name:
                    akt_manschaft.players.remove(akt_manschaft.players[i])
                    tempbool = True
            for i in range(0, len(akt_manschaft.torwart)):
                if akt_manschaft.torwart[i].name == name:
                    akt_manschaft.torwart.remove(akt_manschaft.torwart[i])
                    tempbool = True
            if tempbool:
                print('Erfolgreich Spieler', name, 'entfernt.')
            else:
                print('Spieler nicht gefunden!')
        elif auswahl == 'h':
            name = input('Spielername?')
            nummer = input('Spielernummer?')
            if len(nummer) > 2:
                print('Ungültige Nummer\nNummer darf maximal 2stellig sein')
            elif nummer in ['A', 'B', 'C', 'D']:
                akt_manschaft.trainer.append(spieler(name, nummer))
                akt_manschaft.trainer.sort(key=operator.attrgetter('number'))
            elif nummer.isnumeric():
                nummer = int(nummer)
                if nummer in [1, 12, 16]:
                    akt_manschaft.torwart.append(spieler(name, nummer))
                    akt_manschaft.torwart.sort(key=operator.attrgetter('number'))
                else:
                    akt_manschaft.players.append(spieler(name, nummer))
                    akt_manschaft.players.sort(key=operator.attrgetter('number'))
            else:
                print('Ungültige Nummer')
    elif cmd == 's':
        pickle.dump(Manschaften, open('save.p', 'wb'))
        pickle.dump(Manschaften_kurz, open('save2.p', 'wb'))
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
            print('{:>3}|{:30}|TW'.format(each.number, each.name))
        for each in akt_manschaft.players:
            print('{:>3}|{:30}|'.format(each.number, each.name))
        print('-'*35)
        for each in akt_manschaft.trainer:
            print('{:>3}|{:30}|'.format(each.number, each.name))
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