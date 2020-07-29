#UPDATE DATA IN Ppt FILES  ����Excel�ļ��е�����
#THEN CREATE PDF FILE  Then ����PDF�ĵ�

[string]$path = "D:\powershell\ppts\"  #Path to Ppt spreadsheets to save to PDF  ���浽pdf��excel���ӱ��·��
[string]$savepath = "D:\powershell\pdfs\"
[string]$dToday = Get-Date -Format "yyyyMMdd"


# $ppFixedFormat = "Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType" -as [type] 
# Get-ChildItem ��һ������ָ��λ�û�ȡ��Ŀ������Ŀ
$PptFiles = Get-ChildItem -Path $path -include *.pptx, *.ppt, *.pps  -recurse 

# Create the Excel application object  ����PptӦ�ó������
# New-Object  ����Microsoft .NET Framework��COM�����ʵ��
$objPpt = New-Object -ComObject Powerpoint.application
# $objPpt.visible = $false   #Do not open individual windows  ���򿪵�������

foreach($wb in $PptFiles) 
{ 
# Path to new PDF with date  �������ڵ��µ�PDF��·��
#Join-Path   ��·������·���ϲ�Ϊһ��·����
 $filepath = Join-Path -Path $savepath -ChildPath ($wb.BaseName + "_" + $dtoday + ".pdf") 
 # Open workbook - 3 refreshes links  �򿪹����� 3��ˢ��
 $presentation = $objPpt.presentations.open($wb.fullname,3)
 
 # Give delay to save  �ӳٱ���
 Start-Sleep -s 5
 
 # ����
 $presentation.Saved = $true 
"saving $filepath" 

 #Export as PDF ����ΪPDF
#  $ppFixedFormat::ppFixedFormatTypePDF
#PPT תpdf��ʽ�̶�ֵ 32
# Word תpdf��ʽ�̶�ֵ 17 
 $presentation.SaveAs($filepath,32)
 $presentation.close() 
} 
$objPpt.Quit()