import pandas as pd
import os
import xlsxwriter

def main():
    #chemin du csv
    csvPath = "F:\\Downloads\\LinkViewer\\Source\\Testcsv.csv"
    #chemin où l'on souhaite créer le dossier parent des dossiers des étudiants
    folderPath = "F:\\Downloads\\LinkViewer\\"
    os.chdir(folderPath)
    #Nom du folder du fichier parent
    nameOfFolder = "TestFolder"
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
    datas = pd.read_csv(csvPath, delimiter=";")
    # Creation du header du fichier excel
    header = []
    for head in datas:
        header.append(head)
    #print(header)
    #creation des dossiers par étudiant, avec le fichier excel
    for student in datas["Eleve"]:
        #retour dans le dossier parent par sécurité
        os.chdir(folderPath)
        #print(os.getcwd())
        #creation du dossier etudiant dans le dossier parent
        if not os.path.exists(student):
            os.mkdir(student)
            #print("Directory ", student, " Created ")
        else:
            #print("Directory ", student, " already exists")
            pass
        #changement du cwd pour se mettre dans le dossier étudiant pour créer et ecrire le excel
        os.chdir(folderPath+"\\"+student)
        #print(os.getcwd())
        #creation du excel et de sa feuille
        workbook = xlsxwriter.Workbook(student+".xlsx")
        worksheet = workbook.add_worksheet()
        #creation des données à écrire
        datasToWrite = []
        for head in header:
            #print([head, datas[datas["Eleve"]==student][head].values[0]])
            datasToWrite.append([head, datas[datas["Eleve"]==student][head].values[0]])
        row = 0
        col = 0
        #ecriture des données sur excel, et fermeture du fichier (à l'ouverture, on écrase le contenu du fichier précédent
        for head, value in datasToWrite:
            #print(head, value)
            worksheet.write(row, col, head)
            worksheet.write(row, col+1, value)
            row += 1
        workbook.close()


if __name__ == "__main__":
    main()