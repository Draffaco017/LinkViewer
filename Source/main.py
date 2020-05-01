import pandas as pd
import os, shutil
import xlsxwriter

def main():
    #chemin du csv
    csvPath = "F:\\Downloads\\LinkViewer\\Source\\Carnet de l'apprenant3_Diego.csv"
    #chemin où l'on souhaite créer le dossier parent des dossiers des étudiants
    folderPath = "F:\\Downloads\\LinkViewer\\"
    # newSessionName : pour rajouter un nom "custom" aux séances
    newSessionName = ['S17premierExos', 'S18SecondExo', 'S19TroisiemeExo', 'S19++ exobonus :)']
    os.chdir(folderPath)
    #Nom du folder du fichier parent
    nameOfFolder = "StudentsFolder"
    if not os.path.exists(nameOfFolder):
        os.mkdir(nameOfFolder)
        #print("Directory ", nameOfFolder, " Created ")
    else:
        #print("Directory ", nameOfFolder, " already exists")
        pass
    #mise à jour du fichier de travail avec le fichier parent en tant que nouveau cwd
    folderPath += nameOfFolder+"\\"
    os.chdir(folderPath)
    #lecture des données
    datas = pd.read_csv(csvPath, delimiter=",")
    #print(datas.head())
    # Creation du header du fichier excel
    header = []
    # j'ai trouvé que le mot clé "déposez" était plus robuste pour récupérer le lien drive du fichier
    for head in datas:
        if "Déposez" in head or "Nom" in head or "Prénom" in head:
            header.append(head)
            #print(head, "\n")
    #print(header)
    # c'est le tableau récapitulatif prof qui stockera la progression de tout les étudiants
    datasExcelRecap = []
    #initialiser la forme
    for i in range(datas[header[0]].shape[0]+1):
        datasExcelRecap.append([0]*(len(header)-1))
    #print(datasExcelRecap)
    #initialiser header prof
    for i in range(len(header)):
        if i == 0:
            datasExcelRecap[0][i] = header[i] + " et " + header[i+1]
        elif i == 1:
            pass
        else:
            if i-2 < len(newSessionName):
                datasExcelRecap[0][i - 1] = header[i][0:4]+" "+newSessionName[i-2]
            else:
                datasExcelRecap[0][i - 1] = header[i][0:4]
    #print(datasExcelRecap)
    # creation des dossiers par étudiant, avec le fichier excel
    currentRowForRecap = 1
    for studentName in datas[header[0]]:#key = "Nom"
        studentFolderName = studentName+datas.loc[datas[header[0]] == studentName, header[1]].values[0]
        #retour dans le dossier parent par sécurité
        os.chdir(folderPath)
        #print(os.getcwd())
        #creation du dossier etudiant dans le dossier parent
        if not os.path.exists(studentFolderName):
            os.mkdir(studentFolderName)
            #print("Directory ", studentFolderName, " Created ")
        else:
            #print("Directory ", studentFolderName, " already exists")
            pass
        #vider tout le dossier avant de faire une autre opération:
        otherNameRequired = 0
        for fileName in os.listdir(folderPath + "\\" + studentFolderName):
            filePath = os.path.join(folderPath + "\\" + studentFolderName, fileName)
            try:
                if os.path.isfile(filePath) or os.path.islink(filePath):
                    os.unlink(filePath)
                elif os.path.isdir(filePath):
                    shutil.rmtree(filePath)
            except Exception as e:
                otherNameRequired += 1
        #changement du cwd pour se mettre dans le dossier étudiant pour créer et ecrire le excel
        os.chdir(folderPath+"\\"+studentFolderName)
        #print(os.getcwd())
        #creation des données à écrire
        datasToWrite = []
        for i in range(len(header)):
            #print([head, datas[datas["Eleve"]==student][head].values[0]])
            if i == 0:
                datasToWrite.append([header[i] + " et " + header[i+1], datas.loc[datas[header[0]] == studentName, header[i]].values[0] + " " + datas.loc[datas[header[0]] == studentName, header[i+1]].values[0]])
                datasExcelRecap[currentRowForRecap][i] = (datas.loc[datas[header[0]] == studentName, header[i]].values[0] + " " + datas.loc[datas[header[0]] == studentName, header[i+1]].values[0])
            elif i == 1:
                pass
            else:
                #le header[i][0:4] permet de récupérer le nom de séance (expl : S19+) au début des headers
                #du coup il vaut mieux ne pas changer le type de typographie dans les futures éditions du GForm (et garder une typographie SXX et le mot "Déposez" dans les headers
                if i-2 < len(newSessionName):
                    datasToWrite.append([header[i][0:4]+" "+newSessionName[i-2], datas.loc[datas[header[0]] == studentName, header[i]].values[0]])
                    datasExcelRecap[currentRowForRecap][i-1] = (datas.loc[datas[header[0]] == studentName, header[i]].values[0])
                else:
                    datasToWrite.append([header[i][0:4], datas.loc[datas[header[0]] == studentName, header[i]].values[0]])
                    datasExcelRecap[currentRowForRecap][i - 1] = (datas.loc[datas[header[0]] == studentName, header[i]].values[0])
        row = 0
        col = 0
        # creation du excel et de sa feuille
        if otherNameRequired == 0:
            workbook = xlsxwriter.Workbook(studentFolderName + ".xlsx")
            worksheet = workbook.add_worksheet()
        else:
            workbook = xlsxwriter.Workbook(studentFolderName + " V" + str(otherNameRequired) + ".xlsx")
            worksheet = workbook.add_worksheet()
        #ecriture des données sur excel, et fermeture du fichier (à l'ouverture, on écrase le contenu du fichier précédent
        for head, value in datasToWrite:
            #print(head, value)
            if ";" in str(value):
                #spliter sur plusieurs colonnes si y'a un ";" (plusieurs versions de fichiers)
                #je ne fais ce split que dans les fichiers étudiants, sinon il y aurait des problèmes de layout dans le fichier récapitulatif...
                urls = value.split(";")
                worksheet.write(row, col, head)
                for i in range(len(urls)):
                    worksheet.write(row, col+i+1, str(urls[i]))
            else:
                worksheet.write(row, col, head)
                worksheet.write(row, col + 1, str(value))
            row += 1
        workbook.close()
        currentRowForRecap += 1
    #écriture du fichier récapitulatif prof
    os.chdir(folderPath)
    #print(os.getcwd())
    #print(datasExcelRecap)
    newNameProfFile = 0
    for fileName in os.listdir(folderPath):
        filePath = os.path.join(folderPath, fileName)
        try:
            if os.path.isfile(filePath) or os.path.islink(filePath):
                os.unlink(filePath)
        except Exception as e:
            newNameProfFile += 1
    if newNameProfFile == 0:
        workbook = xlsxwriter.Workbook("RecapProfProduction.xlsx")
    else:
        workbook = xlsxwriter.Workbook("Recap Prof Production V"+str(newNameProfFile)+".xlsx")
    worksheet1 = workbook.add_worksheet("Header == colonne")
    worksheet2 = workbook.add_worksheet("Header == ligne")
    for i in range(len([x[0] for x in datasExcelRecap])):
        # parcours des lignes
        for j in range(len(datasExcelRecap[0][:])):
            #parcours des colonnes
            worksheet1.write(i, j, str(datasExcelRecap[i][j])) #header = ligne
            worksheet2.write(j, i, str(datasExcelRecap[i][j])) #hearder == colonne
    workbook.close()


if __name__ == "__main__":
    main()