[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic");

$cn = [Microsoft.VisualBasic.interaction]::inputbox('输入你的电脑名称', '电脑名字', 'fwx');