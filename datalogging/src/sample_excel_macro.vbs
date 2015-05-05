Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open("C:\Users\mueller\git\data-apps\datalogging\src\SPY200400_ww16.xlsm")

objExcel.Application.Visible = True


objExcel.Application.Run "SPY200400_ww16.xlsm!Log_data2"  




WScript.Echo "Should be logging data now."
WScript.Quit