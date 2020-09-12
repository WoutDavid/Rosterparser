import xlrd as xlrd
from datetime import datetime
import tkinter as Tk
import re as re
import os as os
from tkinter.filedialog import askopenfilename

def runApp():
    window = Tk.Tk()
    window.title("Uurrooster Filter")
    window.geometry('900x200')
    def openXL():
        global fileName
        intermediate = askopenfilename(filetypes=[("Excel files", ".xlsx .xls")])
        fileSplit = intermediate.split(".")
        fileName=fileSplit[0] + "." + fileSplit[-1]
        fileLabel.config(text=fileName)
        fileButton.config(fg="red")

    fileButton = Tk.Button(window,fg="green", text="Select your excel file: ", command=openXL)  
    fileLabel=Tk.Label(window, text="No file selected")
    fileButton.grid(row=0, sticky="W")
    fileLabel.grid(row=0, column=1)
    newFileLabel=Tk.Label(window,text="Insert new filename: ")
    newFileLabel.grid(row=1, sticky="W")
    newFileNameEntry=Tk.Entry(window, width=10)
    newFileNameEntry.grid(row=1, column=1)
    selected = Tk.IntVar()
    firstRadio = Tk.Radiobutton(window,text="First", value=1, variable=selected)
    secondRadio = Tk.Radiobutton(window,text="Second", value=2, variable=selected)
    thirdRadio = Tk.Radiobutton(window,text='Third', value=3, variable=selected)
    fourthRadio = Tk.Radiobutton(window,text='Fourth', value=4, variable=selected)
    radioLabel=Tk.Label(window, text="Which sheet?: ")
    radioLabel.grid(row=2, sticky="W")
    firstRadio.grid(row=2, column=1)
    secondRadio.grid(row=2, column=2)
    thirdRadio.grid(row=2, column=3)
    fourthRadio.grid(row=2, column=4)
    nameLabel = Tk.Label(window, text="What is your name as written on the roster?")
    nameEntry=Tk.Entry(window)
    nameLabel.grid(row=3, sticky="W")
    nameEntry.grid(row=3, column=2)

    def parseXLS():

        if fileName == "":
            errorMsg.config(text="Not a valid filename.")
            return
        if newFileNameEntry.get() == "":
            errorMsg.config(text="Not a valid new fileName.")
            return 
        if nameEntry.get() == "":
            errorMsg.config(text="You forgot to input your name")
            return
        if selected.get()==0:
            errorMsg.config(text="No sheet number indicated.")
            return
        try:
            wb = xlrd.open_workbook(fileName)
        except: 
            errorMsg.config(text="Given file not found.")
            return
        ws = wb.sheet_by_index(selected.get()-1)
        if ".txt" not in newFileNameEntry.get():
            newFileName = newFileNameEntry.get() + ".txt"
        else:
            newFileName = newFileNameEntry.get()
        with open(newFileName, 'w') as newFile:
            for i in range(1,63):
                    voormiddag = ws.cell(i,2)
                    namiddag = ws.cell(i, 3)
                    dag = ws.cell(i, 1)
                    if voormiddag.value!=None:
                        if nameEntry.get().casefold() in str(voormiddag.value).casefold():
                            newFile.write(datetime(*xlrd.xldate_as_tuple(dag.value, wb.datemode)).strftime('%d/%m/%Y') + " Voormiddag" + "\n") 
                    if namiddag.value!=None:
                        if nameEntry.get().casefold() in str(namiddag.value).casefold():
                            newFile.write(datetime(*xlrd.xldate_as_tuple(dag.value, wb.datemode)).strftime('%d/%m/%Y') + " Namiddag" + "\n")
        if os.stat(newFileName).st_size == 0:
            errorMsg.config(text="No name matches were found.")
            return
        def openFile():
            os.startfile(newFileName)
        fileURL = Tk.Button(window, text="Click here to view your file", fg="green", command=openFile)
        fileURL.grid(column=2, row=4)
        

    startBtn = Tk.Button(window, text="Get shifts", bg=("orange"), command=parseXLS)
    startBtn.grid(column=1,row=4)

    errorMsg = Tk.Label(window, text="", fg="red")
    errorMsg.grid(column=1, row=5)
    window.mainloop()           

runApp()