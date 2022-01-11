Set ExcelObj = CreateObject("Excel.Application")
ExcelObj.Visible = true
Set ExcelConfigFile = ExcelObj.Workbooks.Open("C:\Users\91877\Desktop\Automation\PracticeDay4Batch4\Documents\Test.xlsx")
Set Sheet = ExcelConfigFile.Worksheets("Sheet1")
TotalRow = Sheet.Range("A1048576").End(-4162).Row 'xlUp = -4162
WScript.Echo TotalRow