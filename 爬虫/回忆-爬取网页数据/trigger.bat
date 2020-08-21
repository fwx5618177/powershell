@echo off && setlocal enabledelayedexpansion

echo 输入你指定的存储位置
echo 例子----D:\codedata\Powershell\爬虫\回忆-爬取网页数据
echo.
set /p "dirPath=输入你的存储路径："
set /p "pageNum=输入想要获取的数据："
powershell -file 爬取数据.ps1 %dirPath% %pageNum%

echo 运行结束

echo.
echo 根据时间生成excel文件
echo.

del tmp.txt
pause