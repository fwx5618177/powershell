
function createExcel(){

    #显示输出
    Write-Host "xxx" -BackgroundColor Red -ForegroundColor Blue;
    Write-Output "xxx" | Out-File -Append -FilePath "xx.txt";
    
    #获取com的excel对象
    $excel = New-Object -ComObject Excel.Application;
    #设置处理属性和警告
    $excel.Visible = $true;
    $excel.DisplayAlerts = $false;

    #打开指定excel
    $wb = $excel.Workbooks.Open("xx");
    
    #添加sheet
    $workbook = $excel.Workbooks.add();

    #选择sheet
    $sheetDF = $workbook.worksheets.Item(1);

    #获取sheet的个数
    $sheetNum = $wb.sheets.count();

    #获取工作簿的个数
    $excel.Workbooks.Count();

    #获取sheet名字
    $sheetName = $excel.Sheets | Select-Object -Property Name;

    #sheet改名并选择
    $workbook.Worksheets.item(1).name="Processes";
    $sheet = $workbook.WorkSheets.Item("Processes");

    #删除sheet
    $excel.Workbooks.Item(1).delete();
    
    #选定有内容的范围
    $range = $sheet.UsedRange;

    #范围内搜索
    $target = $range.Find("12");
    $range.FindNext("12");
    $range.FindPrevious("12");

    #获取目标行的行、列
    $target.select();
    $target.row;
    $target.column;

    #获取属性的com权柄
    $lineStyle = "microsoft.office.interop.excel.xlLineStyle" -as [type];
    $colorIndex = "microsoft.office.interop.excel.xlColorIndex" -as [type];
    $borderWeight = "microsoft.office.interop.excel.xlBorderWeight" -as [type]
    $chartType = "microsoft.office.interop.excel.xlChartType" -as [type]

    #单元格属性的设置
    $sheet.Cells.item(1,2).font.bold = $true;
    $sheet.Cells.item(1,2).borders.LineStyle = $lineStyle::xlDashDot;
    $sheet.Cells.item(1,2).borders.ColorIndex = $colorIndex::xlColorIndexAutomatic;
    $sheet.Cells.item(1,2).borders.Weight = $borderWeight::xlMedium;
    
    #单元格内部内容属性
    $sheet.Cells.item(1,2).font.ColorIndex = 52;
    $sheet.Cells.item(1,2).interior.colorindex = 20;

    #单元格内容设置
    $sheet.Cells.item(1,1) = "Name of Process";
    $sheet.Cells.item(1,2) = "Size";

    #合并单元格
    $mergeCells = $sheet.Range("A2:A3");
    $mergeCells.select();
    $mergeCells.MergeCells = $true;
    $ma = $sheet.range("A2").mergeArea();
    #返回内部数量
    $ma.count();
    $sheet.Cells.item(1,2).text;
    
    #自动列宽
    $range.EntireColumn.AutoFit();

    #另存文档
    $workbook.saveas("path:/xx.txt");
    
    #关闭工作簿
    $excel.Workbooks.Close();

    #关闭excel
    $workbook.application.quit();

}

#获取进程的数据
$processes=Get-Process

#调用函数
createExcel;

