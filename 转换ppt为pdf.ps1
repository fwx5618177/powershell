#UPDATE DATA IN Ppt FILES  更新Excel文件中的数据
#THEN CREATE PDF FILE  Then 建立PDF文档

[string]$path = "D:\powershell\ppts\"  #Path to Ppt spreadsheets to save to PDF  保存到pdf的excel电子表格路径
[string]$savepath = "D:\powershell\pdfs\"
[string]$dToday = Get-Date -Format "yyyyMMdd"


# $ppFixedFormat = "Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType" -as [type] 
# Get-ChildItem 在一个或多个指定位置获取项目和子项目
$PptFiles = Get-ChildItem -Path $path -include *.pptx, *.ppt, *.pps  -recurse 

# Create the Excel application object  创建Ppt应用程序队象
# New-Object  创建Microsoft .NET Framework或COM对象的实例
$objPpt = New-Object -ComObject Powerpoint.application
# $objPpt.visible = $false   #Do not open individual windows  不打开单个窗口

foreach($wb in $PptFiles) 
{ 
# Path to new PDF with date  带有日期的新的PDF的路径
#Join-Path   将路径和子路径合并为一条路径。
 $filepath = Join-Path -Path $savepath -ChildPath ($wb.BaseName + "_" + $dtoday + ".pdf") 
 # Open workbook - 3 refreshes links  打开工作簿 3秒刷新
 $presentation = $objPpt.presentations.open($wb.fullname,3)
 
 # Give delay to save  延迟保存
 Start-Sleep -s 5
 
 # 保存
 $presentation.Saved = $true 
"saving $filepath" 

 #Export as PDF 导出为PDF
#  $ppFixedFormat::ppFixedFormatTypePDF
#PPT 转pdf格式固定值 32
# Word 转pdf格式固定值 17 
 $presentation.SaveAs($filepath,32)
 $presentation.close() 
} 
$objPpt.Quit()