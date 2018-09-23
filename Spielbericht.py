import PyPDF2
import openpyxl
import operator
import pickle
import easygui
import os

version = '''Version 2.1'''
title = 'Spielbericht ' + version


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
                if rawText[i].isalpha():
                    spieler_name += rawText[i]
                self.trainer.append(spieler(spieler_name, chr(ord('A') + numTrainer)))
                numTrainer += 1
            elif (rawText[i].isalpha() and not rawText[i + 1].isupper()):
                debug = rawText[i - 2:]
                if rawText[i - 2].isnumeric():
                    number = int(rawText[i - 2:i])
                else:
                    if not rawText[i - 1].isnumeric():
                        easygui.msgbox('Manschaft Fehlerhaft')
                        number = 100
                    else:
                        number = int(rawText[i - 1])
                spieler_name = ''
                while (rawText[i].isalpha() and not rawText[i + 1].isupper()) or rawText[i] == ' ':
                    spieler_name += rawText[i]
                    i += 1
                self.players.append(spieler(spieler_name, number))
            i += 1
        for i in range(0, len(self.players) - 1):
            if self.players[i].number > self.players[i + 1].number:
                for u in range(0, i + 1):
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
        if number not in [1, 12, 16] or not number.isalpha():
            if oldNumber in [1, 12, 16]:
                self.torwart.remove(player)
                self.players.append(player)
            elif str(oldNumber).isalpha():
                self.trainer.remove(player)
                self.players.append(player)
            self.players.sort(key=operator.attrgetter('number'))
        elif number.isalpha():
            if oldNumber in [1, 12, 16]:
                self.torwart.remove(player)
                self.trainer.append(player)
            elif str(oldNumber).isnumeric():
                self.players.remove(player)
                self.trainer.append(player)
            self.trainer.sort(key=operator.attrgetter('number'))
        else:
            if str(oldNumber).isaplha():
                self.trainer.remove(player)
                self.torwart.append(player)
            elif str(oldNumber).isnumeric() and oldNumber not in [1, 12, 16]:
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
            akt = easygui.buttonbox('Wollen sie die Manschaft ' + str(name) + ' aktualisieren?', title=title,
                                    choices=['Yes', 'No'])
            if akt == 'Yes':
                del Manschaften[name]
                Manschaften.update({name: manschaft()})
                Manschaften[name].read(name, file)
        else:
            Manschaften.update({name: manschaft()})
            Manschaften[name].read(name, file)
            Manschaften_kurz.update({easygui.enterbox('Kürzel für Manschaft: ' + str(name), title=title): name})
    else:
        Manschaften.update({name: manschaft()})
        Manschaften[name].read(name, file)
        Manschaften_kurz.update({easygui.enterbox('Kürzel für Manschaft: ' + str(name), title=title): name})

    pos_Gast = rawText.find('Gast: ')
    pos_Ende = rawText.find('Nr.Name', pos_Gast)
    name = rawText[pos_Gast + 6:pos_Ende]
    if name in Manschaften.keys():
        akt = easygui.buttonbox('Wollen sie die Manschaft ' + str(name) + ' aktualisieren?', title=title,
                                choices=['Yes', 'No'])
        if akt == 'Yes':
            del Manschaften[name]
            Manschaften.update({name: manschaft()})
            Manschaften[name].read(name, file)
    else:
        while True:
            kuerzel = easygui.enterbox('Kürzel für Manschaft: ' + str(name), title=title)
            if kuerzel in Manschaften_kurz:
                easygui.msgbox('Kürzel schon vorhanden!')
            else:
                Manschaften_kurz.update({kuerzel: name})
                Manschaften.update({name: manschaft()})
                Manschaften[name].read(name, file)
                break


def fileSchreiben():
    global Manschaften, Manschaften_kurz
    heim = easygui.choicebox('Heimmanschaft?', title=title, choices=list(Manschaften_kurz.keys()))
    kurzHeim = heim
    heim = Manschaften_kurz[heim]
    gast = easygui.choicebox('Gastmanschaft?', title=title, choices=list(Manschaften_kurz.keys()))
    gast = Manschaften_kurz[gast]
    heimManschaft = Manschaften[heim]
    gastManschaft = Manschaften[gast]
    datum = easygui.enterbox('Datum?', title=title)
    wb = openpyxl.load_workbook('Mannschaftsliste_MUSTER.xlsx')
    sheet = wb[wb.sheetnames[0]]
    sheet.title = kurzHeim + ' am ' + datum
    sheet['A1'] = heimManschaft.spielklasse
    sheet['A3'] = heimManschaft.name
    sheet['F3'] = gastManschaft.name
    sheet['A2'] = 'Aufstellung vom ' + datum
    for i in range(0, len(heimManschaft.torwart)):
        sheet['B' + str(7 + i)] = heimManschaft.torwart[i].name
        sheet['A' + str(7 + i)] = heimManschaft.torwart[i].number
        sheet['D' + str(7 + i)] = 'TW'
    for i in range(0, len(heimManschaft.players)):
        sheet['B' + str(10 + i)] = heimManschaft.players[i].name
        sheet['A' + str(10 + i)] = heimManschaft.players[i].number
    for i in range(0, len(heimManschaft.trainer)):
        sheet['B' + str(27 + i)] = heimManschaft.trainer[i].name
    for i in range(0, len(gastManschaft.torwart)):
        sheet['G' + str(7 + i)] = gastManschaft.torwart[i].name
        sheet['F' + str(7 + i)] = gastManschaft.torwart[i].number
        sheet['I' + str(7 + i)] = 'TW'
    for i in range(0, len(gastManschaft.players)):
        sheet['G' + str(10 + i)] = gastManschaft.players[i].name
        sheet['F' + str(10 + i)] = gastManschaft.players[i].number
    for i in range(0, len(gastManschaft.trainer)):
        sheet['G' + str(27 + i)] = gastManschaft.trainer[i].name
    date = datum.split('.')
    datei = easygui.enterbox('Dateiname:', title=title, default=kurzHeim + date[-1][2:] + date[1] + date[0])
    try:
        if datei[-5:0] != '.xlsx':
            datei += '.xlsx'
    except:
        datei += '.xlsx'
    wb.save(datei)


print('Spielbericht', version)
try:
    Manschaften = pickle.load(open('save.p', 'rb'))
    Manschaften_kurz = pickle.load(open('save2.p', 'rb'))
except:
    easygui.msgbox('Manschaften konnten nicht geladen werden!')
    Manschaften = {}
    Manschaften_kurz = {}

choices = ['Quit', 'Bogen kreiren', 'Datei lesen', 'Kürzelliste anzeigen', 'Spielerliste einer Manschaft anzeigen',
           'Speichern',
           'Manschaft editieren', 'Manschaft löschen', 'Manschaft hinzufügen    (Manuell)', 'Kürzel ändern']

while True:
    if len(Manschaften_kurz) == 0:
        cmd = easygui.choicebox('Was wollen Sie machen?', title=title,
                                choices=['Quit', 'Datei lesen', 'Manschaft hinzufügen    (Manuell)'])
    else:
        cmd = easygui.choicebox('Was wollen Sie machen?', title=title, choices=choices)
    if cmd == 'Quit':
        easygui.msgbox('Goodbye!', title=title)
        break
    elif cmd == 'Kürzel ändern':
        akt_manschaft = easygui.choicebox('Von welcher Manschaft soll das kürzel geändert werden?', title=title,
                                          choices=list(Manschaften_kurz.keys()))
        k = akt_manschaft
        akt_manschaft = Manschaften_kurz[akt_manschaft]
        new_kuerzel = easygui.enterbox('Neues Kürzel?', title=title)
        if new_kuerzel in Manschaften_kurz.keys():
            easygui.msgbox('Kürzel schon vorhanden! Kürzel wird nicht geändert!', title=title)
        else:
            del Manschaften_kurz[k]
            Manschaften_kurz.update({new_kuerzel: akt_manschaft})
            easygui.msgbox(
                'Erfolgreich Kürzel der Manschschaft ' + str(akt_manschaft) + ' von ' + str(k) + ' auf ' + str(
                    new_kuerzel))
    elif cmd == 'Manschaft hinzufügen    (Manuell)':
        name = easygui.enterbox('Manschaftname?', title=title)
        kuerzel = easygui.enterbox('Kürzel?', title=title)
        if kuerzel in Manschaften_kurz.keys():
            easygui.msgbox('Kürzel schon vergeben!', title=title)
            continue
        spielklasse = easygui.enterbox('Spielklasse?', title=title)
        temp = easygui.buttonbox(
            'Wollen sie wirklich die Manschaft ' + name + ' mit dem Kürzel ' + kuerzel + ' in der Spielklasse ' + spielklasse + 'kreieren',
            title=title, choices=['Yes', 'No'])
        if temp == 'Yes':
            Manschaften_kurz.update({kuerzel: name})
            Manschaften.update({name: manschaft()})
            Manschaften[name].name = name
            Manschaften[name].spielklasse = spielklasse
            easygui.msgbox('Manschaft erfolgreich hinzugefügt!')
            pickle.dump(Manschaften, open('save.p', 'wb'))
            pickle.dump(Manschaften_kurz, open('save2.p', 'wb'))

    elif cmd == 'Manschaft löschen':
        akt_manschaft = easygui.choicebox('Welche Manschaft soll gelöscht werden?', title=title,
                                          choices=list(Manschaften_kurz.keys()))
        kuerzel = akt_manschaft
        akt_manschaft = Manschaften_kurz[akt_manschaft]
        temp = easygui.buttonbox('Wollen sie wirklich die Manschaft ' + str(akt_manschaft) + ' löschen?',
                                 title=title, choices=['Yes', 'No'])
        if temp == 'Yes':
            del Manschaften_kurz[kuerzel]
            del Manschaften[akt_manschaft]
    elif cmd == 'Manschaft editieren':
        akt_manschaft = easygui.choicebox('Welche Manschaft soll editiert werden?', title=title,
                                          choices=list(Manschaften_kurz.keys()))
        akt_manschaft = Manschaften_kurz[akt_manschaft]
        akt_manschaft = Manschaften[akt_manschaft]
        auswahl = easygui.buttonbox('Was willst du machen?', title=title,
                                    choices=['Spieler löschen', 'Spielernummer ändern',
                                             'Spieler hinzufügen', 'Spielername ändern', 'Quit'])
        if auswahl == 'Spielernummer ändern':
            name = easygui.enterbox('Wie heißt der Spieler?', title=title)
            new_number = easygui.integerbox('Welche nummer soll der Spieler bekommen?', title=title, lowerbound=1,
                                            upperbound=99)
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
                if akt_manschaft.trainer[i].name == name:
                    akt_manschaft.changeNumberPlayer(akt_manschaft.trainer[i], new_number)
                    tempbool = True
            if tempbool:
                easygui.msgbox('Erfolgreich Nummer von ' + name + ' auf ' + str(new_number) + ' gesetzt.', title=title)
            else:
                easygui.msgbox('Spieler nicht gefunden!', title=title)
        elif auswahl == 'Spielername ändern':
            name = easygui.enterbox('Wie heißt der Spieler?', title=title)
            new_name = easygui.enterbox('Welche Namen soll der Spieler bekommen?', title=title)
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
                easygui.msgbox('Erfolgreich Nummer von ' + name + 'auf' + new_name + ' gesetzt.', title=title)
            else:
                easygui.msgbox('Spieler nicht gefunden!', title=title)
        elif auswahl == 'Spieler löschen':
            name = easygui.enterbox('Wie heißt der Spieler?', title=title)
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
                easygui.msgbox('Erfolgreich Spieler' + name + 'entfernt.', title=title)
            else:
                easygui.msgbox('Spieler nicht gefunden!', title=title)
        elif auswahl == 'Spieler hinzufügen':
            name = easygui.enterbox('Spielername?', title=title)
            nummer = easygui.enterbox('Spielernummer?', title=title)
            if nummer in ['A', 'B', 'C', 'D']:
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
                easygui.msgbox('Ungültige Nummer! Vorgang abgebrochen!', title=title)
    elif cmd == 'Speichern':
        pickle.dump(Manschaften, open('save.p', 'wb'))
        pickle.dump(Manschaften_kurz, open('save2.p', 'wb'))
    elif cmd == 'Spielerliste einer Manschaft anzeigen':
        akt_manschaft = easygui.choicebox('Welche Manschaft soll gelöscht werden?', title=title,
                                          choices=list(Manschaften_kurz.keys()))
        akt_manschaft = Manschaften_kurz[akt_manschaft]
        akt_manschaft = Manschaften[akt_manschaft]
        ausgabe = ''
        for each in akt_manschaft.torwart:
            ausgabe += '{:>3}|{:30}|TW'.format(each.number, each.name)
            ausgabe += '\n'
        for each in akt_manschaft.players:
            ausgabe += '{:>3}|{:30}|'.format(each.number, each.name)
            ausgabe += '\n'
        ausgabe += '-' * 35
        ausgabe += '\n'
        for each in akt_manschaft.trainer:
            ausgabe += '{:>3}|{:30}|'.format(each.number, each.name)
            ausgabe += '\n'
        easygui.msgbox(ausgabe, title=title)
    elif cmd == 'Kürzelliste anzeigen':
        ausgabe = ''
        for each in Manschaften_kurz.keys():
            ausgabe += each
            ausgabe += '\n'
            ausgabe += Manschaften_kurz[each]
            ausgabe += '\n'
            ausgabe += Manschaften[Manschaften_kurz[each]].spielklasse
            ausgabe += '\n\n'
        easygui.msgbox(ausgabe, title=title)
    elif cmd == 'Bogen kreiren':
        fileSchreiben()
    elif cmd == 'Datei lesen':
        dirs = os.listdir()
        for i in dirs[:]:
            try:
                if not i[-4:] == '.pdf':
                    dirs.remove(i)
            except:
                dirs.remove(i)
        datei = easygui.choicebox('Dateiname?', title=title, choices=dirs)
        if datei[-4:] != '.pdf':
            datei += '.pdf'
        fileRead(datei)
        pickle.dump(Manschaften, open('save.p', 'wb'))
        pickle.dump(Manschaften_kurz, open('save2.p', 'wb'))

pickle.dump(Manschaften, open('save.p', 'wb'))
pickle.dump(Manschaften_kurz, open('save2.p', 'wb'))
