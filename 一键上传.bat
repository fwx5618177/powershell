@echo off
set /p str=Input :

git status
git add .
git commit -am "%str%"