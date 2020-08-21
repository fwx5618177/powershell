@echo off

set /p "DIRPATHSTR=文件存放地址："
set /p "SEARCHSTR=搜索关键字："
set /p "PAGES=抓取数量：（例子，1-3页写为1..3）"

powershell -file 爬图.ps1 %DIRPATHSTR% %SEARCHSTR% %PAGES%