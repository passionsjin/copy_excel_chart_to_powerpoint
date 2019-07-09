import win32com.client as win32
from win32com.client import Dispatch
from win32com.client import constants
import os

# https://stackoverflow.com/questions/11110752/export-charts-from-excel-as-images-using-python
# https://stackoverflow.com/questions/32639900/charts-from-excel-to-powerpoint-with-python
xlApp = Dispatch('Excel.Application')
pptApp = Dispatch('PowerPoint.Application')
workbook = xlApp.Workbooks.Open('C:\\Users\\Park\\Desktop\\private_project\\POC_chart_copy_to_ppt\\data_chart.xlsx')

presentation = pptApp.Presentations.Add(True)
xlApp.Sheets("Sheet1").Select()

xlSheet1 = xlApp.Sheets(1)
# for ws in workbook.Worksheets:
for chart in xlSheet1.ChartObjects():
    print(chart.name)
    chart.Activate()
    chart.Copy()
    print(constants.__dicts__)
    # https://www.thespreadsheetguru.com/blog/2014/3/17/copy-paste-an-excel-range-into-powerpoint-with-vba
    # https://docs.microsoft.com/ko-kr/office/vba/api/PowerPoint.PpSlideLayout
    Slide = presentation.Slides.Add(presentation.Slides.Count + 1, 12)
    Slide.Shapes.PasteSpecial(11)

    textbox = Slide.Shapes.AddTextbox(1, 100, 100, 200, 300)
    textbox.TextFrame.TextRange.Text = str(chart.Chart.ChartTitle.Text)

presentation.SaveAs("C:\\Users\\Park\\Desktop\\private_project\\POC_chart_copy_to_ppt\\ppt_test.pptx")
presentation.Close()
pptApp.Quit()

# CLOSE

xlApp.ActiveWorkbook.Close()

# presentation.Quit()
# xlApp.Quit()

