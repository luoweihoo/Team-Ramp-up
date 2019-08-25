# @File    :   MakeupTestList.py
# @Time    :   2019/08/25 22:51:28
# @Author  :   Wei Luo 
# @Version :   1.0
# @Contact :   luoweihoo@yahoo.com
# @Desc    :   Generate the make-up test list of a workshop

import openpyxl, sys
from openpyxl.utils import get_column_letter

trackList = 'Team Ramp-up Tracking.xlsx'

# Check if the track list is in the current folder; if yes open it
print('Opening workbook...')
try:
    wb = openpyxl.load_workbook(trackList)
except FileNotFoundError:
    print(f"The track list {trackList} doesn't exist, please check!")

# Build the workshop list
sheet = wb.active
workshopList = []
colNum = 1
for cell in sheet[2]:
    if '.' in str(cell.value):
        cellString = str(cell.value).split('.')
        colName = get_column_letter(colNum)
        cellString.append(colName)
        workshopList.append(cellString)
    colNum = colNum + 1    
    
# Ask user to enter the No. of the workshop
rawInput = input("Please enter the number of the workshop!\n e.g. 1, 2..., 'q' to quit program:")
if str(rawInput).upper() == 'Q':
    sys.exit()

# Confirm the chosen workshop
workshopNum = str(rawInput)
for workshop in workshopList:
    if workshopNum == workshop[0]:
        print(f"The workshop you chose is {workshop[1].lstrip()}.")
        rawInput = input("Please confirm the action? Y or N:")
        if str(rawInput).upper() == 'N':
            sys.exit()
        else:
            break

# Record the column name in which the workshop is in
workshopCol = str(workshop[2])

# The chosen workshop
# print(workshopNum)
# print(workshop[1])
# print(workshop[2])

# Generate the make-up test list and save it to a seperate .xlsx file
makeupList = str(workshop[0]) + '_' + 'MakeupTestList.xlsx'
listWb = openpyxl.Workbook()
listSheet = listWb.active
listRow = 1
for row in range(4, sheet.max_row + 1):
    if sheet[workshopCol + str(row)].value == None:
        pass
    elif sheet[workshopCol + str(row)].value == 'X' or sheet[workshopCol + str(row)].value == 'X*' or int(sheet[workshopCol + str(row)].value) < 60:
        listSheet['A' + str(listRow)].value = sheet['A' + str(row)].value
        listSheet['B' + str(listRow)].value = sheet['B' + str(row)].value
        listRow = listRow + 1

listWb.save(makeupList)
print('Processing is finished!')
