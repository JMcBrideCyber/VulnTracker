#"C:\\Users\\JLMcB\\Documents\\TestingChecklists"
import enum
from tkinter.tix import COLUMN
import xml.etree.ElementTree as elt
import os
import glob
from os.path import isfile, join
import xlsxwriter
import VulnTrackerInterface
from VulnTrackerInterface import *

#Get a name for the excel file
fileName = VulnTrackerInterface.startWork.fileName
print(VulnTrackerInterface.startWork.fileName)

#Opening a new Excel workbook with the filename given by user
wb = xlsxwriter.Workbook(fileName + ".xlsx")

#Adding a worksheet, default name
ws = wb.add_worksheet()

#Formatting for text
#Headers
header_format = wb.add_format()
header_format.set_font_size(20)

#Normal Cells
cell_format = wb.add_format()
cell_format.set_font_size(12)

#Vulnerability Totals, color for clarity
#Cat 1
cat1_format = wb.add_format()
cat1_format.set_font_size(12)
cat1_format.set_bg_color('red')
cat1_format.set_bold()

#Cat 2
cat2_format = wb.add_format()
cat2_format.set_font_size(12)
cat2_format.set_bg_color('orange')
cat2_format.set_bold()

#Cat 3
cat3_format = wb.add_format()
cat3_format.set_font_size(12)
cat3_format.set_bg_color('yellow')
cat3_format.set_bold()

#Setting Excel column and row values, as well as a counter variable
row = 0
col = 0
counter = 0

#Counters to track total vulnerabilities accross all checklists
totalCat1 = 0
totalCat2 = 0
totalCat3 = 0

tree = elt.ElementTree()

#List of lists that will hold the data from each checklist
info = []

#Change this to browse for file
folder = glob.glob("C:\\Users\\JLMcB\\Documents\\TestingChecklists\\*.ckl")

#Iterates through each checklist in a folder
for checklist in folder:

	#Counters to keep track of totals for each level of vulnerability
	cat1 = 0
	cat2 = 0
	cat3 = 0

	#Writing the name of the checklist to the Excel sheetl, only occurs when a new ckl is being added
	col = 0
	ws.write(row, col, (os.path.basename(checklist).strip(".ckl")), header_format)
	row += 1
	
	#Parsing the data from a single checklist
	tree.parse(checklist)

	#List that will hold the information gathered from a checklist
	nextInfo = []

	#Navigate to the correct layer of the hierarchy and set the root at that level
	root = tree.findall("STIGS/iSTIG/VULN")

	#loops through each tag below "VULN" and look for the desired tags to gather info
	for cklData in root:
		
		col = 0

		#Get the status of the given check NR/NA/NAF/O
		status = cklData.find("STATUS")

		#If check if open, continue proccessing
		if status.text == "Open":

			#Grabbing the information in "STIG_DATA" tag
			stigData = cklData.findall("STIG_DATA")

			#Looking at the info in "STIG_DATA" tags
			for data in stigData:
		
				#Switch statement to get the information within fields that have data we want Vuln_Num/Rule_Title/Severity
				match data[0].text:
					case "Vuln_Num":

						#Getting information from field 
						ws.write(row, col, data[1].text.strip("\n"), cell_format)
						col += 1
						#Appending gathered info to the nextInfo list to be added to the list of lists "info"
						continue
					case "Rule_Title":

						#Getting information from field
						ws.write(row, col, data[1].text.strip("\n"), cell_format)
						row += 1
						#Appending gathered info to the nextInfo list to be added to the list of lists "info"
						continue
					case "Severity":

						#Getting information from field
						ws.write(row, col, data[1].text.strip("\n"), cell_format)
						col += 1
						#Appending gathered info to the nextInfo list to be added to the list of lists "info"
						if data[1].text.strip("\n") == "low":
							cat3 += 1
							totalCat3 += 1
						if data[1].text.strip("\n") == "medium":
							cat2 += 1
							totalCat2 += 1
						if data[1].text.strip("\n") == "high":
							cat1 += 1
							totalCat1 += 1
						continue
						
	ws.write(row, col, "Cat 1s: " + str(cat1), cat1_format)
	col += 1
	ws.write(row, col, "Cat 2s: " + str(cat2), cat2_format)
	col += 1
	ws.write(row, col, "Cat 3s: " + str(cat3), cat3_format)
	row += 1			

#Formatting for totals
col = 0
ws.write(row,col,"")
row += 1
ws.write(row, col, "Total", cat1_format)
col += 1 
ws.write(row, col, "Total", cat2_format)
col += 1 
ws.write(row, col, "Total", cat3_format)

col = 0
row += 1

ws.write(row, col, "Cat 1s: " + str(totalCat1), cat1_format)
col += 1
ws.write(row, col, "Cat 1s: " + str(totalCat2), cat2_format)
col += 1
ws.write(row, col, "Cat 1s: " + str(totalCat3), cat3_format)

#Close workbook
wb.close()
VulnTrackerInterface.window.destroy


