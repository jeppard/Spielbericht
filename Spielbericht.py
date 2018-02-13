import PyPDF2
import openpyxl
version = '''Version 0.1\nby Jonas Eppard'''

class manschaft():
    def __init__(self, name, file):
        self.name = name
        pdfFile = open(file, 'rb')
        pdfReader = PyPDF2.PdfFileReader(pdfFile)
        pageObj = pdfReader.getPage(1)
        rawText = pageObj.extractText()
        pos_Manschaft = rawText.find(name)
        self.players = []
        self.trainer = []
        i = 0
        rawText = rawText[pos_Manschaft + len(name) + 80:]

        while True:
            if rawText[i:i + 4] == 'Gast' or rawText[i:i + 8] == 'Handball':
                break
            if (rawText[i] in ['A', 'B', 'C', 'D'] and
                    rawText[i + 1].isupper() and rawText[i + 2].islower() and
                    not (rawText[i:i + 6] == 'DGast:' or rawText[i:i + 9] == 'DHandball')):
                i += 1
                spieler = ''
                while (rawText[i].isalpha() and not rawText[i + 1].isupper()) or rawText[i] == ' ':
                    spieler += rawText[i]
                    i += 1
                self.trainer.append(spieler)
            elif (rawText[i].isalpha() and not rawText[i + 1].isupper()):
                spieler = ''
                while (rawText[i].isalpha() and not rawText[i + 1].isupper()) or rawText[i] == ' ':
                    spieler += rawText[i]
                    i += 1
                self.players.append(spieler)
            i += 1

print(version)
heim = input('Heimmanschaft?')
if heim == '':
    heim = 'HSG Böblingen/Sindelfingen'
    print('Heim automatisch auf', heim, 'gesetzt')
heimFile = input('Datei(Heim)')
gast = input('Gastmanschaft?')
gastFile = input('Datei(Gast)')
heimManschaft = manschaft(heim, heimFile)
gastManschaft = manschaft(gast, gastFile)
print(heimManschaft.players)
print(gastManschaft.players)
wb = openpyxl.load_workbook('Mannschaftsliste_MUSTER.xlsx')
sheet = wb[wb.sheetnames[0]]
sheet['A3'] = heimManschaft.name
sheet['F3'] = gastManschaft.name
for i in range(0, len(heimManschaft.players)):
    sheet['B'+str(7+i)] = heimManschaft.players[i]
for i in range(0, len(heimManschaft.trainer)):
    sheet['B'+str(27+i)] = heimManschaft.trainer[i]
for i in range(0, len(gastManschaft.players)):
    sheet['G'+str(7+i)] = gastManschaft.players[i]
for i in range(0, len(gastManschaft.trainer)):
    sheet['G'+str(27+i)] = gastManschaft.trainer[i]
datei = input('Dateiname:')
try:
    if datei[-5:0] != '.xlsx':
        datei += '.xlsx'
except:
    datei += '.xlsx'
wb.save(datei)