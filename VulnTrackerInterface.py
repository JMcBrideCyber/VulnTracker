#import VulnTracker
from cgitb import text
from email import header
from fileinput import filename
from sre_parse import State
import tkinter as tk
from tkinter import *
from tkinter import filedialog
from numpy import size
import os
import VulnTracker

def browseFiles():
    browseFiles.fileName = filedialog.askdirectory(initialdir = "/", title = "Choose a file")
    fileReadback.config(state = NORMAL)
    fileReadback.insert(1.0, browseFiles.fileName)
    fileReadback.config(state = DISABLED)

def startWork():
    wbName = wbNameBox.get(1.0, "end-1c")
    sheetName = sheetNameBox.get(1.0, "end-1c")
    fileName = browseFiles.fileName
    os.system("VulnTracker.py")
    #window.destroy()
    
    


window = Tk()
window.title("VulnTracker")
window.geometry("500x200")

#Header that tells the user what to do
header = Label(window, text = "Select a directory of checklists to process.", font = (("Calibri"), 15))
header.place(anchor=CENTER, relx = .5, rely = .1)

#Tells user to enter a name for the Excel workbook
wbNamePrompt = Label(window, text = "Workbook Name: ", font = (("Calibri"), 15))
wbNamePrompt.place(anchor = CENTER, relx = .165, rely = .45)

#Text box for user to enter a workbook name
wbNameBox = Text(height = 1, width = 35)
wbNameBox.place(anchor = CENTER, relx = .65, rely = .45)

#Button to open file explorer
browseFilesButton = Button(window, text = "Browse Files", command = browseFiles, height = 1, width = 20)
browseFilesButton.place(anchor=CENTER, relx = .16, rely = .25)

#Text box that reads back the directory selected
fileReadback = Text(height = 1, width = 35, state = DISABLED)
fileReadback.place(anchor = CENTER, relx = .65, rely = .25)

#Tells user to enter a name for the Excel worksheet
sheetNamePrompt = Label(window, text = "Worksheet Name: ", font = (("Calibri"), 15))
sheetNamePrompt.place(anchor = CENTER, relx = .17, rely = .65)

#Text box for user to enter a worksheet name
sheetNameBox = Text(height = 1, width = 35)
sheetNameBox.place(anchor = CENTER, relx = .65, rely = .65)

#Button to get all the inputs and begin processing
startButton = Button(window, text = "Start", height = 2, width = 20, command = startWork)
startButton.place(anchor = CENTER, relx = .5, rely = .85)

#Keeps window open
window.mainloop()
