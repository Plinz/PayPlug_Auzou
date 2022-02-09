import PySimpleGUI as sg
import csv

layout = [[sg.Text('Fichier CSV PayPlug')],
          [sg.Input(key='input'), sg.FileBrowse(target='input', file_types=(("CSV Files", "*.csv"),))],
          [sg.Text('Fichier à générer')],
          [sg.Input(key='output') , sg.SaveAs(target='output', file_types=(("CSV Files", "*.csv"),))],
          [sg.OK(), sg.Cancel()]]

window = sg.Window('PayPlug data transformer', layout)
event, values = window.read()

def removeChars(text, chars):
    return ''.join(c for c in text if c not in chars)
    
def parseInputFile(csvInput, csvOutput):
    with open(csvOutput, 'w', newline='') as csvfileoutput:
        writer = csv.writer(csvfileoutput, delimiter=';')
        with open(csvInput) as csvfileinput:
            reader = list(csv.reader(csvfileinput, delimiter=','))
            i = 1
            while (i < len(reader) and reader[i][0]):
                date = removeChars(reader[i][1].split(' ')[0], ' "')
                ref = removeChars(reader[i][0], ' "')
                val = removeChars(reader[i][5], ' -+"')
                type = removeChars(reader[i][3], ' "')
                if ('Paiement' in type):
                    listLib = removeChars(reader[i][10], '"').split(',')
                    order = listLib[3].split(':')[1].split('}')[0].strip()
                    client = listLib[0].split(':')[1].split('}')[0].strip()
                    lib = 'FA' + str(int(order)-10).zfill(6) + ',Client:' + client.zfill(5) + ',Order:' + order.zfill(6)
                    writer.writerow(['149', date, '51710000', '', ref, lib, val, ''])
                    writer.writerow(['149', date, '41190000', '4000', ref, lib, '', val])
                elif ('Remboursement' in type):
                    listLib = removeChars(reader[i][10], '"').split(',')
                    client = listLib[0].split(':')[1].split('}')[0].strip()
                    lib = '#' + removeChars(reader[i][4], '"').zfill(8) + ',Client:' + client.zfill(5)
                    writer.writerow(['149', date, '51710000', '', ref, lib, '', val])
                    writer.writerow(['149', date, '41190000', '4000', ref, lib, val, ''])
                elif ('Facture' in type):
                    lib = removeChars(reader[i][4], '"')
                    writer.writerow(['149', date, '51710000', '', ref, lib, '', val])
                    writer.writerow(['149', date, '41190000', '401387', ref, lib, val, ''])
                else:
                    print("[ERROR] [" + csvInput + "] La ligne " + str(i+1) + " n'a pas été reconnu. Valeur de la ligne :" + type)
                i += 1

csvInput = values['input']
csvOutput = values['output']
parseInputFile(csvInput, csvOutput)
window.close()