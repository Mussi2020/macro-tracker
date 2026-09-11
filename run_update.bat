@echo off
rem Macro Dashboard - update entry point (used by 任务计划程序 & 立即更新.bat)
rem 依赖装在 .pylibs 里，任意 Python 3.9+ 都能跑，无需重建虚拟环境。
cd /d "%~dp0"

set "PY="
for %%P in (
  "D:\PY\python.exe"
  "%USERPROFILE%\.workbuddy\binaries\python\versions\3.13.12\python.exe"
  "C:\Python313\python.exe"
) do (
  if not defined PY if exist %%P set "PY=%%~P"
)
if not defined PY set "PY=python"

"%PY%" "%~dp0update_macro.py"
exit /b %ERRORLEVEL%
