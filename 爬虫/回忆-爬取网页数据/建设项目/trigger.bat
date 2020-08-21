@echo off && setlocal enabledelayedexpansion


echo 输入你指定的存储位置
echo 例子----D:\codedata\Powershell\爬虫\回忆-爬取网页数据
echo.
set /p "dirPath=输入你的存储路径："
set /p "PageNum=输入你需要爬取的页码：（例 1..3）"

powershell -file 建设项目爬虫.ps1 %dirPath% %PageNum%


echo.
echo 根据时间生成excel文件
echo.


pause