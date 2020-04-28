import pandas as pd
import sys
from tkinter import *
import webbrowser
import os, sys


class MainGui(Frame):
    def __init__(self, parent):
        Frame.__init__(self, parent)
        self.parent = parent
        self.initUi()

    def initUi(self):
        self.parent.title("list of students and links")
        self.pack(fill=BOTH, expand=True)
        scrollbarStudent = Scrollbar(self)

        self.students = Listbox(self, yscrollcommand=scrollbarStudent.set)
        scrollbarStudent.pack(side=LEFT, fill=Y)
        self.students.pack(side=LEFT, expand=True, fill=BOTH)

        scrollbarLinks = Scrollbar(self)

        self.links = Listbox(self, yscrollcommand=scrollbarLinks.set)
        scrollbarLinks.pack(side=RIGHT, fill=Y)
        self.links.pack(side=RIGHT, expand=True, fill=BOTH)



    def initLists(self, students):
        for student in students:
            self.students.insert(END, student)


def onSelectStudent(event, datas, links):
    try:
        w = event.widget
        index = int(w.curselection()[0])
        value = w.get(index)
        # print(index, value)
        links.delete(0, 'end')
        compt = 0
        for link in datas[datas["Eleve"] == value].values[0]:
            if compt == 0:
                pass
                compt += 1
            else:
                links.insert(END, "lien vers prépa "+ str(compt) + " : "+ link)
    except IndexError:
        pass


def onSelectLink(event):
    try:
        w = event.widget
        index = int(w.curselection()[0])
        value = w.get(index)
        #print(index, value)
        webbrowser.open(findURL(value)[0])
    except IndexError:
        pass

def findURL(string):
    # findall() has been used
    # with valid conditions for urls in string
    return re.findall('http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\(\), ]|(?:%[0-9a-fA-F][0-9a-fA-F]))+', string)

def main():
    fileName = "Testcsv.csv"
    datas = pd.read_csv(fileName, delimiter=";")
    mainWindow = Tk()
    app = MainGui(mainWindow)
    app.initLists(datas["Eleve"])
    app.students.bind('<<ListboxSelect>>', lambda event, datas=datas, links = app.links : onSelectStudent(event, datas, links))
    app.links.bind('<<ListboxSelect>>', onSelectLink)
    mainWindow.mainloop()


if __name__=="__main__":
    main()