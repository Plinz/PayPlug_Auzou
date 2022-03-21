import PySimpleGUI as sg
import csv, os

def removeChars(text, chars):
    return ''.join(c for c in text if c not in chars)
    
def popupError(errorText):
    error = sg.Window(windowName, [[sg.Text(errorText)], [sg.Column([[sg.Button("Continuer"), sg.Button("Arrêter")]], element_justification='c', expand_x=True)]], modal = True)
    eventError, valuesError = error.read()
    error.close()
    return eventError


def parseInputFile(csvInput, csvOutput):
    with open(csvOutput, 'w', newline='') as csvfileoutput:
        writer = csv.writer(csvfileoutput, delimiter=';')
        with open(csvInput) as csvfileinput:
            reader = list(csv.reader(csvfileinput, delimiter=','))
            i = 1
            while (i < len(reader) and reader[i][0]):
                if len(reader[i]) > 11 :
                    date = removeChars(reader[i][1].split(' ')[0], ' "')
                    date = date[8:] + date[5:7] + date[2:4]
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
                        codePaiement = reader[i][4].strip().split('#')[-1]
                        client = listLib[0].split(':')[1].split('}')[0].strip()
                        lib = 'Remboursement #' + codePaiement + ',Client:' + client.zfill(5)
                        writer.writerow(['149', date, '51710000', '', ref, lib, '', val])
                        writer.writerow(['149', date, '41190000', '4000', ref, lib, val, ''])
                    elif ('Opposition' in type):
                        codePaiement = reader[i][4].strip().split('#')[-1]
                        lib = 'Opposition #' + codePaiement
                        writer.writerow(['149', date, '51710000', '', ref, lib, '', val])
                        writer.writerow(['149', date, '41190000', '4000', ref, lib, val, ''])
                    elif ('Facture' in type):
                        lib = removeChars(reader[i][4], '"')
                        writer.writerow(['149', date, '51710000', '', ref, lib, '', val])
                        writer.writerow(['149', date, '40110000', '401387', ref, lib, val, ''])
                    else:
                        eventError = popupError("Le Type (" + type + ") de la ligne " + str(i+1) + " n'a pas été reconnu.")
                        if eventError != "Continuer" :
                            return False
                else :
                    eventError = popupError("La ligne " + str(i+1) + " n'a pas pu être transformée, car le nombre de colonnes est insuffisant")
                    if eventError != "Continuer" :
                        return False
                i += 1
    return True
    
def runScript(values):
    csvInput = values['input']
    csvOutput = values['output']
    isCompleted = parseInputFile(csvInput, csvOutput)
    endText = "Fin de l'exécution" if isCompleted else "L'exécution a été interrompu"
    window = sg.Window(windowName, [[sg.Text(endText)], [sg.Button('Ouvrir le dossier'), sg.Button('Ouvrir le fichier'), sg.Button('Fermer')]], modal = True)
    event, values = window.read()
    window.close()
    if event == 'Ouvrir le dossier':
        path = os.path.dirname(os.path.abspath(csvOutput))
        os.startfile(path)
    elif event == 'Ouvrir le fichier' :
        path = os.path.realpath(csvOutput)
        os.startfile(path)

windowName = 'PayPlug Transformer - AUZOU EDITIONS'
layout = [[sg.Text(text='AUZOU EDITIONS', font='_ 20 bold', justification='c', expand_x=True)],
          [sg.Text('Fichier CSV PayPlug')],
          [sg.Input(key='input'), sg.FileBrowse("Parcourir", target='input', file_types=(("Fichiers CSV", "*.csv"),))],
          [sg.Text('Fichier à générer')],
          [sg.Input(key='output') , sg.SaveAs("Enregistrer sous", target='output', file_types=(("Fichiers Texte", "*.txt"),))],
          [sg.Column([[sg.OK(), sg.Cancel("Annuler")]], element_justification='c', expand_x=True)]]

window = sg.Window(windowName, layout, margins=(0, 0))
event, values = window.read()
if event == 'OK' :
    runScript(values)
