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
    # Creation du header des fichiers excels
    header = []
    # j'ai trouvé que le mot clé "Déposez" était plus robuste pour récupérer le lien drive du fichier
    for head in datas:
        if "Déposez" in head or "Nom" in head or "Prénom" in head or "N° Equipe" in head:
            header.append(head)
            #print(head, "\n")
    #print(header)
    # enlever les espaces des noms et prénoms
    for name in datas[header[1]].values:
        #print(datas.loc[datas[header[1]] == name, header[1]].values[0])
        datas.loc[datas[header[1]] == name, header[1]] = name.replace(" ", "")
        #print(datas.loc[datas[header[1]] == name.replace(" ", ""), header[1]].values[0])
    for firstName in datas[header[2]].values:
        #print(datas.loc[datas[header[2]] == firstName, header[2]].values[0])
        datas.loc[datas[header[2]] == firstName, header[2]] = firstName.replace(" ", "")
        #print(datas.loc[datas[header[2]] == firstName.replace(" ", ""), header[2]].values[0])
    # c'est le tableau récapitulatif prof qui stockera la progression de tout les étudiants
    datasExcelRecap = []
    #initialiser la forme
    for i in range(datas[header[0]].shape[0]+1):
        datasExcelRecap.append([0]*(len(header)))
    #print(datasExcelRecap)
    #initialiser header prof
    for i in range(len(header)):
        if i == 0:#equipe
            datasExcelRecap[0][i] = header[i]
        elif i == 1:#nom
            datasExcelRecap[0][i] = header[i]
        elif i == 2:#prénom
            datasExcelRecap[0][i] = header[i]
        else:
            if i-3 < len(newSessionName):
                datasExcelRecap[0][i] = header[i][0:4]+" "+newSessionName[i-3]
            else:
                datasExcelRecap[0][i] = header[i][0:4]
    #print(datasExcelRecap)
    # creation des dossiers par étudiant, avec le fichier excel
    currentRowForRecap = 1
    for studentName in datas[header[1]]:#key = "Nom"
        studentFolderName = studentName+datas.loc[datas[header[1]] == studentName, header[2]].values[0]
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
            if i == 0:#equipe pour fichier prof, tout pour le student
                datasToWrite.append([header[i+1] + ", " + header[i + 2] + " et " + header[i],
                                     datas.loc[datas[header[1]] == studentName, header[i + 1]].values[0] + " " +
                                     datas.loc[datas[header[1]] == studentName, header[i + 2]].values[0] + " " +
                                    "Equipe " + str(datas.loc[datas[header[1]] == studentName, header[i]].values[0])])
                datasExcelRecap[currentRowForRecap][i] = str(datas.loc[datas[header[1]] == studentName, header[i]].values[0])
            elif i == 1:#nom pour fichier prof
                datasExcelRecap[currentRowForRecap][i] = datas.loc[datas[header[1]] == studentName, header[i]].values[0]
            elif i == 2:#prenom pour fichier prof
                datasExcelRecap[currentRowForRecap][i] = datas.loc[datas[header[1]] == studentName, header[i]].values[0]
            else:
                #le header[i][0:4] permet de récupérer le nom de séance (expl : S19+) au début des headers
                #du coup il vaut mieux ne pas changer le type de typographie dans les futures éditions du GForm (et garder une typographie SXX et le mot "Déposez" dans les headers
                if i-3 < len(newSessionName):
                    datasToWrite.append([header[i][0:4]+" "+newSessionName[i-3], datas.loc[datas[header[1]] == studentName, header[i]].values[0]])
                    datasExcelRecap[currentRowForRecap][i] = (datas.loc[datas[header[1]] == studentName, header[i]].values[0])
                else:
                    datasToWrite.append([header[i][0:4], datas.loc[datas[header[1]] == studentName, header[i]].values[0]])
                    datasExcelRecap[currentRowForRecap][i] = (datas.loc[datas[header[1]] == studentName, header[i]].values[0])
        row = 0
        col = 0
        # creation du excel et de sa feuille
        if otherNameRequired == 0:
            workbook = xlsxwriter.Workbook(studentFolderName + ".xlsx")
            worksheetStudent = workbook.add_worksheet()
        else:
            workbook = xlsxwriter.Workbook(studentFolderName + " V" + str(otherNameRequired) + ".xlsx")
            worksheetStudent = workbook.add_worksheet()
        #ecriture des données sur excel, et fermeture du fichier (à l'ouverture, on écrase le contenu du fichier précédent
        for head, value in datasToWrite:
            #print(head, value)
            if ";" in str(value):
                #spliter sur plusieurs colonnes si y'a un ";" (plusieurs versions de fichiers)
                #je ne fais ce split que dans les fichiers étudiants, sinon il y aurait des problèmes de layout dans le fichier récapitulatif...
                urls = value.split(";")
                for i in range(len(urls)):
                    worksheetStudent.write(row, col, head + "V" + str(i+1))
                    worksheetStudent.write(row, col+1, str(urls[i]))
                    row += 1
            else:
                worksheetStudent.write(row, col, head)
                worksheetStudent.write(row, col + 1, str(value))
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
    worksheet2 = workbook.add_worksheet("Header == colonne, ligne split")
    worksheet3 = workbook.add_worksheet("Header == ligne")
    for i in range(len([x[0] for x in datasExcelRecap])):
        # parcours des lignes
        for j in range(len(datasExcelRecap[0][:])):
            #parcours des colonnes
            worksheet1.write(i, j, str(datasExcelRecap[i][j])) #header == ligne
            worksheet3.write(j, i, str(datasExcelRecap[i][j]))  # header == colonne
            #décommenter le if else si dessous et commenter la ligne
            # worksheet3.write(j, i, str(datasExcelRecap[i][j])) #juste au dessus
            #afin d'eviter dans header == colonne que la première ligne soit le nom de l'équipe
            # if j == 0:
            #     pass
            # else:
            #     worksheet3.write(j-1, i, str(datasExcelRecap[i][j])) #header == colonne
    rowWorksheet2 = 0
    columnWorksheet2 = 0
    #écriture du header
    for i in range(3):
        worksheet2.write(rowWorksheet2, i, datasExcelRecap[0][i])
    worksheet2.write(rowWorksheet2, 3, "Séance")
    worksheet2.write(rowWorksheet2, 4, "Lien")
    rowWorksheet2 += 1
    for student in (datasExcelRecap[1:]):
        #print(student)
        currentCourseColumn = 3
        for urls in student[3:]:
            #print(urls)
            for url in str(urls).split(";"):
                for i in range(3):#ecrire equipe, nom et prenom
                    worksheet2.write(rowWorksheet2, i, student[i])
                #ecrire le nom de la séance
                worksheet2.write(rowWorksheet2, 3, datasExcelRecap[0][currentCourseColumn])
                #enfin, écrire le lien
                worksheet2.write(rowWorksheet2, 4, str(url))
                rowWorksheet2 += 1
            currentCourseColumn += 1
    workbook.close()


if __name__ == "__main__":
    main()