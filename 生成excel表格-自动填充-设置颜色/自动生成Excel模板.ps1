$processes=Get-Process
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$excel.DisplayAlerts = $false;
$workbook = $excel.Workbooks.add()
$sheet = $workbook.worksheets.Item(1)
#$workbook.Worksheets.item(3).delete()
#$workbook.Worksheets.item(2).delete()
$workbook.Worksheets.item(1).name="Processes"
$sheet = $workbook.WorkSheets.Item("Processes")
$x = 2

$lineStyle = "microsoft.office.interop.excel.xlLineStyle" -as [type]
$colorIndex = "microsoft.office.interop.excel.xlColorIndex" -as [type]
$borderWeight = "microsoft.office.interop.excel.xlBorderWeight" -as [type]
$chartType = "microsoft.office.interop.excel.xlChartType" -as [type]

for($b = 1 ; $b -le 2 ; $b++)
{
 $sheet.cells.item(1,$b).font.bold = $true
 $sheet.cells.item(1,$b).borders.LineStyle = $lineStyle::xlDashDot
 $sheet.cells.item(1,$b).borders.ColorIndex = $colorIndex::xlColorIndexAutomatic
# $sheet.cells.item(1,$b).borders.Weight = $borderWeight::xlMedium
}

$sheet.cells.item(1,1) = "Name of Process"
$sheet.cells.item(1,2) = "Working Set Size"

foreach($process in $processes)
{
 $sheet.cells.item($x, 1) = $process.name
 $sheet.cells.item($x, 1).font.ColorIndex = 52 #只有这个起作用，
 # $sheet.cells.item($x, 1).font.Color = 40
    #$sheet.cells.item($x, 1).interior.color = 22
  $sheet.cells.item($x, 1).interior.colorindex = 20#只有这个起作用，
 $sheet.cells.item($x,2) = $process.workingSet
 $x++
 if ($x -gt 10) {
 break
 }
} #end foreach

$range = $sheet.usedRange

$MergeCells = $sheet.Range("A2:A5")
$MergeCells.Select()
$MergeCells.MergeCells = $true 


$range.EntireColumn.AutoFit() | out-null

$ma = $sheet.range("A2").mergeArea()
write-host $ma.count()
 write-host "merge aaa= " $sheet.cells.item(2,2).text -ForegroundColor red

for ($i=0;$i -lt $ma.count();$i++) {
   write-host "merge aera for A2= " $sheet.cells.item(2+$i,2).text -ForegroundColor yellow
}

$currentDir = Split-Path -Parent $MyInvocation.MyCommand.Definition;


$workbook.saveas("$currentDir\测试.xlsx");
$workbook.application.quit();