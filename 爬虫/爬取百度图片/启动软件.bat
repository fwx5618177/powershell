@echo off

set /p "DIRPATHSTR=�ļ���ŵ�ַ��"
set /p "SEARCHSTR=�����ؼ��֣�"
set /p "PAGES=ץȡ�����������ӣ�1-3ҳдΪ1..3��"

powershell -file ��ͼ.ps1 %DIRPATHSTR% %SEARCHSTR% %PAGES%