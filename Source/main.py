import pandas as pd
import os
import xlsxwriter

def main():
    #chemin du csv
    csvPath = "F:\\Downloads\\LinkViewer\\Source\\Carnet de l'apprenant3_Diego.csv"
    #chemin où l'on souhaite créer le dossier parent des dossiers des étudiants
    folderPath = "F:\\Downloads\\LinkViewer\\"
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
        if "Déposez" in head or "Nom" in head or "Prénom" in head :
            header.append(head)
            #print(head, "\n")
    #print(header)
    #creation des dossiers par étudiant, avec le fichier excel
    datasExcelRecap = []#c'est le tableau qui stockera le excel récapitulatif pour chaque étudiant
    for studentName in datas[header[0]]:#key = "Nom"
        studentFolderName = studentName+datas.loc[datas[header[0]]==studentName, header[1]].values[0]
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
        #changement du cwd pour se mettre dans le dossier étudiant pour créer et ecrire le excel
        os.chdir(folderPath+"\\"+studentFolderName)
        #print(os.getcwd())
        #creation du excel et de sa feuille
        workbook = xlsxwriter.Workbook(studentFolderName+".xlsx")
        worksheet = workbook.add_worksheet()
        #creation des données à écrire
        datasToWrite = []
        #newSessionName : pour rajouter un nom "custom" aux séances
        newSessionName = ['S17premierExos', 'S18SecondExo', 'S19TroisiemeExo', 'S19++Bonus:)']
        for i in range(len(header)):
            #print([head, datas[datas["Eleve"]==student][head].values[0]])
            if i == 0:
                datasToWrite.append([header[i]+" et "+header[i+1], datas.loc[datas[header[0]] == studentName, header[i]].values[0]+" "+datas.loc[datas[header[0]] == studentName, header[i+1]].values[0]])
            elif i == 1:
                pass
            else:
                #le header[i][0:3] permet de récupérer le nom de séance (expl : S19+) au début des headers
                #du coup il vaut mieux ne pas changer le type de typographie dans les futures éditions du GForm (et garder une typographie SXX et le mot "Déposez" dans les headers
                datasToWrite.append([header[i][0:3]+" "+newSessionName[i-2], datas.loc[datas[header[0]] == studentName, header[i]].values[0]])
        row = 0
        col = 0
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


if __name__ == "__main__":
    main()