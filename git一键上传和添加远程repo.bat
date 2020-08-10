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
set /p str=文件更改信息：
git add .
git commit -am "%str%"

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
pause
goto start

:s0
exit