$processes=Get-Process
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$workbook = $excel.Workbooks.add()
$sheet = $workbook.worksheets.Item(1)
$workbook.Worksheets.item(1).name="Processes"
$sheet = $workbook.WorkSheets.Item("Processes")