import win32com.client as win32
from win32com.client import Dispatch
from win32com.client import constants
import os

# https://stackoverflow.com/questions/11110752/export-charts-from-excel-as-images-using-python
# https://stackoverflow.com/questions/32639900/charts-from-excel-to-powerpoint-with-python
xlApp = Dispatch('Excel.Application')
pptApp = Dispatch('PowerPoint.Application')

base_path = os.path.dirname(os.path.abspath(__file__))
xl_path = os.path.join(base_path, 'data_chart.xlsx')
ppt_path = os.path.join(base_path, 'target_ppt.pptx')

workbook = xlApp.Workbooks.Open(xl_path)

# presentation = pptApp.Presentations.Add(True)
presentation = pptApp.Presentations.Open(ppt_path)
xlApp.Sheets("Sheet1").Select()

xlSheet1 = xlApp.Sheets(1)
# for ws in workbook.Worksheets:
count = 1
for chart in xlSheet1.ChartObjects():
    print(chart.name)
    chart.Activate()
    chart.Copy()
    print(constants.__dicts__)
    # https://www.thespreadsheetguru.com/blog/2014/3/17/copy-paste-an-excel-range-into-powerpoint-with-vba
    # https://docs.microsoft.com/ko-kr/office/vba/api/PowerPoint.PpSlideLayout
    # https://docs.microsoft.com/en-us/office/vba/api/powerpoint.shapes.pastespecial
    # Slide = presentation.Slides.Add(presentation.Slides.Count + 1, 12)
    Slide = presentation.Slides.Item(1)
    pchart = Slide.Shapes.PasteSpecial(11)
    # Left, Top, Width, Height 로 지정 가능. pchart.Height = 500
    pchart.Top = 400 * (count - 1)

    textbox = Slide.Shapes.AddTextbox(1, 100, 100, 200, 300)
    textbox.TextFrame.TextRange.Text = str(chart.Chart.ChartTitle.Text)

    count += 1

# presentation.SaveAs(ppt_path)
# presentation.Close()
# pptApp.Quit()

# CLOSE

xlApp.ActiveWorkbook.Close()

# presentation.Quit()
# xlApp.Quit()

