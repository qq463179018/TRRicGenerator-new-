@echo off

REM Command is last param
setlocal
set valname=%1
:loop
if {%1}=={} goto found
set script=%1
shift
goto :loop
:found

REM Test for script
IF NOT EXIST "%SKitRoot%\bin\%script%.pl" GOTO ERROR

REM Chain script with correct perl environment
perl -I"%SKitRoot%\lib\perl" "%SKitRoot%\bin\%script%.pl" %*
endlocal
exit /B %ERRORLEVEL%

REM On error
:ERROR
  echo [ERROR] Invalid SKitRoot or script missing.
  echo         Unable to locate %SKitRoot%\bin\%script%.pl
