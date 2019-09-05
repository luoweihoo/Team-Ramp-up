# @File    :   create_summary.py
# @Time    :   2019/09/05 15:48:28
# @Author  :   Wei Luo 
# @Version :   1.0
# @Contact :   luoweihoo@yahoo.com
# @Desc    :   Create the summary into a PPT slides deck

import openpyxl, sys, trfunction
from openpyxl.styles import Font
from openpyxl.styles.colors import BLACK
from openpyxl.styles.colors import RED
from pptx import Presentation


# trackList = 'Team Ramp-up Tracking.xlsx'
# # Check if the track list is in the current folder; if yes open it
# wb = trfunction.checkTrackList(trackList)
# # Build the workshop list
# workshopList = trfunction.buildWorkshopList(wb)
# # Ask user to enter the No. of the workshop
# workshopNum, workshopCol = trfunction.chooseWorkshop(workshopList)
# # Open the test result of the selected workshop
# testResult, isMakeupTest = trfunction.openTestResult(workshopNum)

# print(testResult)

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[0])
slide.shapes.title.text = "Yo, python!"
slide.placeholders[1].text = "yes it's really awesome"
prs.save('yoPython.pptx')

# print('Processing is finished!')
