from tkinter import *
from tkinter import ttk, messagebox
from tkinter.filedialog import askopenfilename, askdirectory
from xlrd import *
import pandas as pd 
import configparser
import os.path


def createConfigfile():
    config = configparser.ConfigParser()
    config['DEFAULTS'] = {'Filename':'',
                            'SaveDirectory':''}
    with open('configuration.ini','w') as configfile:
        config.write(configfile)

def readConfigfile():
    if not os.path.isfile('configuration.ini'):
        createConfigfile()

    config = configparser.ConfigParser()
    config.read('configuration.ini')
    print(config['DEFAULTS']['Filename'])
    localfilename = config['DEFAULTS']['Filename']
    localdirectory = config['DEFAULTS']['SaveDirectory']
    return localfilename, localdirectory

def saveConfigfile():
    localfilename = myfileName.get()
    localdirectoryname = myDirectoryName.get()
    config = configparser.ConfigParser()

    config['DEFAULTS'] = {'Filename':localfilename,
                            'SaveDirectory':localdirectoryname}
    with open('configuration.ini','w') as configfile:
        config.write(configfile)

def selectFile():
    try:
        value = askopenfilename()
        myfileName.set(value)
    except ValueError:
        pass

def selectFileDirectory():
    try:
        value = askdirectory()
        myDirectoryName.set(value)
    except ValueError:
        pass        

def splitExcelFile():
    inputfile = myfileName.get()     
    xl = open_workbook(inputfile)    
    mysheetnames = xl.sheet_names()
    path = myDirectoryName.get()+'/'

    for name in mysheetnames:
        writer = pd.ExcelWriter(path+name+'.xlsx')
        parsing = pd.ExcelFile(inputfile).parse(sheet_name=name)
        parsing.to_excel(writer,name)
        writer.save()
    saveConfigfile()
    messagebox.showinfo("Splitten erfolgreich", "Die Exceldatei wurde aufgesplittet.")


root = Tk()
root.title("Excelfile-Splitter")

mainframe = ttk.Frame(root, padding="3 3 12 12")
mainframe.grid(column=0, row=0, sticky=(N, W, E, S))
root.columnconfigure(0, weight=1)
root.rowconfigure(0, weight=1)

myfileName = StringVar()
myDirectoryName = StringVar()

tempfileName, tempdirectory = readConfigfile()
myfileName.set(tempfileName)
myDirectoryName.set(tempdirectory)


# Select file to split
#Label
ttk.Label(mainframe,text="Dateiname").grid(column=1, row=1,sticky=(W))
#Entry
fileName_entry = ttk.Entry(mainframe, width = 20, textvariable = myfileName)
fileName_entry.grid(column=2, row=1,sticky=(W,E))
ttk.Button(mainframe, text="Datei auswählen", command=selectFile).grid(column=3, row=1, sticky=W)

#Select target directory
#Label
ttk.Label(mainframe,text="Ausgabeverzeichnis").grid(column=1, row=2,sticky=(E))
fileName_entry = ttk.Entry(mainframe, width = 20, textvariable = myDirectoryName)
fileName_entry.grid(column=2, row=2,sticky=(W,E))
ttk.Button(mainframe, text="Verzeichnis auswählen", command=selectFileDirectory).grid(column=3, row=2, sticky=W)

ttk.Button(mainframe, text="Exceldatei splitten", command=splitExcelFile).grid(column=1,columnspan=3,row=3, sticky=(E,W))


#for child in mainframe.winfo_children(): child.grid_configure(padx=5, pady=5)

root.mainloop()

