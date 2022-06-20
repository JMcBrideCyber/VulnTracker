#"C:\\Users\\JLMcB\\Documents\\TestingChecklists"
import enum
from tkinter.tix import COLUMN
import xml.etree.ElementTree as elt
import os
import glob
from os.path import isfile, join
import xlsxwriter

tree = elt.ElementTree()

info = []
checklists = []

folder = glob.glob("C:\\Users\\JLMcB\\Documents\\TestingChecklists\\*.ckl")

print (folder)

for checklist in folder:
	
	print(os.path.basename(checklist).strip(".ckl"))
	checklistName = (os.path.basename(checklist).strip(".ckl"))
	checklists.append(checklistName)
	tree.parse(checklist)

	nextInfo = []

	root = tree.findall("STIGS/iSTIG/VULN")

	for cklData in root:
		status = cklData.find("STATUS")
		if status.text == "Open":
			stigData = cklData.findall("STIG_DATA")
			for data in stigData:
				match data[0].text:
					case "Vuln_Num":
						vulnNum = data[1].text.strip("\n")
						print(vulnNum)
						nextInfo.append(vulnNum)
						continue
					case "Rule_Title":
						ruleTitle = data[1].text.strip("\n")
						print(ruleTitle)
						nextInfo.append(ruleTitle)
						continue
					case "Severity":
						severity = data[1].text.strip("\n")
						print(severity)
						nextInfo.append(severity)
						continue
	
	info.append(nextInfo)



	wb = xlsxwriter.Workbook("Test.xlsx")
	ws = wb.add_worksheet()
	row = 0
	col = 0

	for line in checklists:
		for line in info:
			for item in line:
				ws.write(row, col, item)
				col += 1
				row += 1

	wb.close()

print(checklists)
#print(info)


