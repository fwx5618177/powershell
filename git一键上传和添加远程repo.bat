@echo off
setlocal enabledelayedexpansion

:start
echo.
echo ѡ�����Ŀ�ģ�
echo.
echo 1.�鿴״̬(status)
echo 2.add, but not commit
echo 3.ȫ���ļ���add + commit
echo 4.���Զ�ֿ̲�
echo 5.���Զ�push��Զ�ֿ̲�
echo 6.ȫ�Զ�push��Ĭ�ϵ�һ���ֿ��master�汾��
echo 0.�˳�
echo.

:select
set /p select=ѡ��

IF /I "%select%"=="1" GOTO s1
IF /I "%select%"=="2" GOTO s2
IF /I "%select%"=="3" GOTO s3
IF /I "%select%"=="4" GOTO s4
IF /I "%select%"=="5" GOTO s5
IF /I "%select%"=="6" GOTO s6
IF /I "%select%"=="0" GOTO s0

ECHO ѡ����Ч������������ 
ECHO.
GOTO select

:s1
echo ��ǰ״̬��
git status
echo.
pause
goto start

:s2
echo add:
set /p str=��ӵľ����ļ�����
git add %str%
echo.
pause
goto start

:s3
echo commit:
set /p str=commit��Ϣ��
git add .
git commit -am "%str%"
echo.
pause
goto start

:s4
echo ����Զ�ֿ̲�
set /p name=����ֿ����֣�
set /p url=����ֿ����ӣ�
git remote add %name% %url%
echo.
pause
goto start

:s5
echo ���Զ��ϴ�
set /p str=�ļ�������Ϣ��
git add .
git commit -am "%str%"

for /f %%i in ('git remote') do (
	echo Զ�ֿ̲����֣�%%i
)
set /p repoName=�ֿ����֣�

for /f "tokens=1,2* delims=\ " %%i in ('git branch') do (
	echo �汾��%%j
)
set /p branch=Ҫ�ϴ��İ汾��

git push %repoName% %branch%

echo.
echo �ϴ����
pause
goto start


:s6
set repoName=
set branch=

echo ȫ�Զ��ϴ�
set /p str=ȫ���ļ�������Ϣ��¼��

::�õ�Ŀ¼����
for /f "tokens=4 delims=\ " %%i in ('dir ^| findstr /c:"%date:~0,10%" ^| findstr /v /c:"\."') do (
	echo %%i>>list.txt
)

::�õ��ļ�����������
for /f %%i in (list.txt) do (
	
	cd %%i

	::��ȡ�����ļ���
	for /f "tokens=4 delims=\ " %%a in ('dir ^| findstr /c:"%date:~0,10%" ^| 	findstr /v /c:"DIR"') do (
		echo %%a >> ..\fileList.txt
	) 
	cd ..
)


::git add .
::����ļ�������
for /f %%i in (list.txt) do (
	git add %%i
	cd %%i
	::��ȡ�����ļ���
	::git commit -am "%str%"
	::����޸���Ϣ
	for /f "tokens=4 delims=\ " %%a in ('dir ^| findstr /c:"%date:~0,10%" ^| findstr /v /c:"DIR"') do (
		echo %%a
		git commit -m "add func: %%a in %date:~0,10%-%time%" %%a
	)
	cd ..
	
)

type list.txt >> UploadLog.txt
echo %date:~0,10%-%time% >> UploadLog.txt
echo. >> UploadLog.txt
type fileList.txt >> UploadLog.txt
echo %date:~0,10%-%time% >> UploadLog.txt
echo. >> UploadLog.txt

::��ȡ����������һ����ʾ

for /f %%i in (fileList.txt) do (
	set /p="%%i,"<nul>>tmplist.txt
)

for /f %%i in (tmplist.txt) do (
	set tmpStr=%%i
)

echo !tmpStr:~,-1!

del list.txt fileList.txt tmplist.txt


git add .
git commit -am "modify: %str% add: !tmpStr:~,-1! in %date:~0,10%-%time% "

for /f %%i in ('git remote') do (
	echo Զ�ֿ̲����֣�%%i
	set repoName=%%i
)

for /f "tokens=1,2* delims=\ " %%i in ('git branch') do (
	echo �汾��%%j
	set branch=%%j
)

echo !repoName!
echo !branch!

git push !repoName! !branch!

echo.
echo �ϴ����
pause > nul
goto start

:s0
exit