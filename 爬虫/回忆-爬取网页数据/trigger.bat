@echo off && setlocal enabledelayedexpansion

echo ������ָ���Ĵ洢λ��
echo ����----D:\codedata\Powershell\����\����-��ȡ��ҳ����
echo.
set /p "dirPath=������Ĵ洢·����"
set /p "pageNum=������Ҫ��ȡ�����ݣ�"
powershell -file ��ȡ����.ps1 %dirPath% %pageNum%

echo ���н���

echo.
echo ����ʱ������excel�ļ�
echo.

del tmp.txt
pause