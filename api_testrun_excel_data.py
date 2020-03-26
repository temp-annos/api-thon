import openpyxl
import os
import json
import pandas as pd

#Load the excel for parameters 
wb = openpyxl.load_workbook('~/Downloads/API.xlsx') 
sheet = wb['Sheet1']

# read vlaues from the cell 
def copyrange(startCol, startRow, endCol, endRow, sheet):
    rangeSelected = []
#Loops through selected Rows
    for i in range(startRow,endRow + 1,1):
#Appends the row to a RowSelected list
        rowSelected = []
        for j in range(startCol,endCol+1,1):
            rowSelected.append(sheet.cell(row = i, column = j).value)
#Adds the RowSelected List and nests inside the rangeSelected
        rangeSelected.append(rowSelected)
    return rangeSelected

# loop to count number of users
for cell in sheet["A"]:
    if cell.value is None:
        print(cell.row)
        break
else:
    print(cell.row + 1)
a=cell.row 
print(a)
rangeSelected = copyrange(1,a,27,a,sheet)

# To check data copied
print(rangeSelected[0][1])
print('This is done 1')
		
# Entering data in Payload
payload = { }

# To check feteched 
print(payload)
print('This is done 2') 

#
payload = json.dumps(payload)
print(payload)
