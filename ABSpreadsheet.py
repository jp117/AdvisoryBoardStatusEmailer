import openpyxl
import os

os.chdir('c:\\users\\John Paradise\\MyPythonScripts\\AdvisoryBoard')

workbook = openpyxl.load_workbook('AdvisoryBoardTracking.xlsx')
newSub = workbook['NewSubmissions']
pendingSub = workbook['PendingReSubmissions']

#Finds all the rows in the sheet to search data for
def nextrow(sheet):
	i = 1
	row = sheet.cell(row=i, column=1).value
	while row != None:
		i += 1
		row = sheet.cell(row=i, column=1).value
	return i

#Makes a list of lists of all jobs that are awaiting submission 
def newSubList(sheet):
	datalistoflist = []
	i = 2
	for i in range (i, nextrow(sheet)):
		data = sheet.cell(row=i, column=14).value
		if data == None:
			datalist = []
			for x in range (1,14):
				datalist.append(sheet.cell(row=i, column=x).value)
			datalistoflist.append(datalist)
	return datalistoflist

#Makes a list of lists of all jobs that are pending and awaiting resubmission
def pendingSubList(sheet):
	datalistoflist = []
	i = 2
	for i in range (i, nextrow(sheet)):
		data = sheet.cell(row=i, column=10).value
		if data == None or data.lower != "yes":
			datalist = []
			for x in range (1,10):
				datalist.append(sheet.cell(row=i, column=x).value)
			datalistoflist.append(datalist)
		return datalistoflist

#Makes the string that will be emailed to the salesmen
def salesmanEmailBody():
	i = len(newSubList(newSub))
	for x in range (x, i)

#Code below accesses any piece of the list of lists that i need.  First is list of lists index number, second is the inside list index number
#print (newSubList(newSub)[0][1])

#this code below for lookig up length of a list.  useful for looping
#print(len(newSubList(newSub)[0]))