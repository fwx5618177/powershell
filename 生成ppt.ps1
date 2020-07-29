Add-type -AssemblyName office
$ppt = New-Object -ComObject Powerpoint.Application;

# 设置可视化
$ppt.Visible = [Microsoft.Office.Core.MsoTriState]::msoCTrue;
#弹窗
$ppt.displayalerts = [Microsoft.Office.Core.MsoTriState]::msoCTrue;

#添加slip
$slip = $ppt.Presentations.Add();

#增加内容
$ppt.presentation.close();
$slip.Application.quit();

$ppt = $null;
[GC]::Collect();