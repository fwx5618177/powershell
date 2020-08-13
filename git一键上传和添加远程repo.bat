@echo off
setlocal enabledelayedexpansion

:start
echo.
echo 选择你的目的：
echo.
echo 1.查看状态(status)
echo 2.add, but not commit
echo 3.全部文件：add + commit
echo 4.添加远程仓库
echo 5.半自动push到远程仓库
echo 6.全自动push（默认第一个仓库和master版本）
echo 0.退出
echo.

:select
set /p select=选择：

IF /I "%select%"=="1" GOTO s1
IF /I "%select%"=="2" GOTO s2
IF /I "%select%"=="3" GOTO s3
IF /I "%select%"=="4" GOTO s4
IF /I "%select%"=="5" GOTO s5
IF /I "%select%"=="6" GOTO s6
IF /I "%select%"=="0" GOTO s0

ECHO 选择无效，请重新输入 
ECHO.
GOTO select

:s1
echo 当前状态：
git status
echo.
pause
goto start

:s2
echo add:
set /p str=添加的具体文件名：
git add %str%
echo.
pause
goto start

:s3
echo commit:
set /p str=commit信息：
git add .
git commit -am "%str%"
echo.
pause
goto start

:s4
echo 增加远程仓库
set /p name=输入仓库名字：
set /p url=输入仓库链接：
git remote add %name% %url%
echo.
pause
goto start

:s5
echo 半自动上传
set /p str=文件更改信息：
git add .
git commit -am "%str%"

for /f %%i in ('git remote') do (
	echo 远程仓库名字：%%i
)
set /p repoName=仓库名字：

for /f "tokens=1,2* delims=\ " %%i in ('git branch') do (
	echo 版本：%%j
)
set /p branch=要上传的版本：

git push %repoName% %branch%

echo.
echo 上传完成
pause
goto start


:s6
set repoName=
set branch=

echo 全自动上传
set /p str=全部文件更改信息记录：

::得到目录数据
for /f "tokens=4 delims=\ " %%i in ('dir ^| findstr /c:"%date:~0,10%" ^| findstr /v /c:"\."') do (
	echo %%i>>list.txt
)

::得到文件的新增数据
for /f %%i in (list.txt) do (
	
	cd %%i

	::获取新增文件名
	for /f "tokens=4 delims=\ " %%a in ('dir ^| findstr /c:"%date:~0,10%" ^| 	findstr /v /c:"DIR"') do (
		echo %%a >> ..\fileList.txt
	) 
	cd ..
)


::git add .
::逐个文件夹增加
for /f %%i in (list.txt) do (
	git add %%i
	cd %%i
	::获取新增文件名
	::git commit -am "%str%"
	::添加修改信息
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

::获取更改名单，一行显示

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
	echo 远程仓库名字：%%i
	set repoName=%%i
)

for /f "tokens=1,2* delims=\ " %%i in ('git branch') do (
	echo 版本：%%j
	set branch=%%j
)

echo !repoName!
echo !branch!

git push !repoName! !branch!

echo.
echo 上传完成
pause > nul
goto start

:s0
exit