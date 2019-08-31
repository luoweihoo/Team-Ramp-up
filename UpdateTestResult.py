# @File    :   UpdateTestResult.py
# @Time    :   2019/08/25 12:27:28
# @Author  :   Wei Luo 
# @Version :   1.0
# @Contact :   luoweihoo@yahoo.com
# @Desc    :   Update the test of each ramp-up session

import openpyxl, sys, trfunction
from openpyxl.styles import Font
from openpyxl.styles.colors import BLACK
from openpyxl.styles.colors import RED

trackList = 'Team Ramp-up Tracking.xlsx'
# Check if the track list is in the current folder; if yes open it
wb = trfunction.checkTrackList(trackList)
# Build the workshop list
workshopList = trfunction.buildWorkshopList(wb)
# Ask user to enter the No. of the workshop
workshopNum, workshopCol = trfunction.chooseWorkshop(workshopList)
# Open the test result of the selected workshop
testResult, isMakeupTest = trfunction.openTestResult(workshopNum)

# print(testResult)
# print(workshopCol)

# Update the test resule into the tracking list
sheet = wb.active
for result in testResult:
    for row in range(4, sheet.max_row + 1):
        if sheet['A' + str(row)].value == result[0]:
            if isMakeupTest == 'N':
                sheet[workshopCol + str(row)].font = Font(color = BLACK)
            else:
                sheet[workshopCol + str(row)].font = Font(color = RED)
            sheet[workshopCol + str(row)].value = result[2]

wb.save(trackList)
print('Processing is finished!')
