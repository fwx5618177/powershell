@echo off && setlocal enabledelayedexpansion


echo ������ָ���Ĵ洢λ��
echo ����----D:\codedata\Powershell\����\����-��ȡ��ҳ����
echo.
set /p "dirPath=������Ĵ洢·����"
set /p "PageNum=��������Ҫ��ȡ��ҳ�룺���� 1..3��"

powershell -file ������Ŀ����.ps1 %dirPath% %PageNum%


echo.
echo ����ʱ������excel�ļ�
echo.


pause