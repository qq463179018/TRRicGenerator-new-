@echo off
echo ************* SKit Environment ************
echo.
set SKitRoot
ftype SKitLet 2>NUL
assoc .skt 2>NUL
set PATH
echo.
echo ************ Perl Environment *************
echo.
call perl "%SKitRoot%\bin\perl_ver.pl" 2>NUL
if NOT ERRORLEVEL = 1 goto op_env
echo Perl Not-installed
:op_env
echo.
echo ********** Operating Environment **********
ver
echo.
echo *************** SKit Version **************
echo.
echo SKit v3.0.0
echo.
echo ************* SKitLets Version ************
echo.
call skitlet_version 2>NUL
if NOT ERRORLEVEL = 1 goto cont1
echo SKitLets Not-supported 
:cont1
echo.
echo ************ SKit Tool Versions ***********
echo.
cscript/nologo "%SKitRoot%\bin\skit_version_int.vbs"

REM po_prepare
:po_prepare
call skit_get_version po_prepare 2>NUL
if NOT ERRORLEVEL = 1 goto dbout_diff
echo po_prepare Not-supported 

REM DBOut_Diff
:dbout_diff
call skit_get_version dbout_diff 2>NUL
if NOT ERRORLEVEL = 1 goto market_summary
echo DBOut_Diff Not-supported 

REM Market Summary
:market_summary
call skit_get_version market_summary 2>NUL
if NOT ERRORLEVEL = 1 goto tracklosttrades
echo MarketSummary Not-supported

REM TrackLostTrades
:tracklosttrades
call skit_get_version track_lost_trades 2>NUL
if NOT ERRORLEVEL = 1 goto market_holidays
echo TrackLostTrades Not-supported

REM MarketHolidays
:market_holidays
call skit_get_version market_holidays 2>NUL
if NOT ERRORLEVEL = 1 goto rv
echo MarketHolidays Not-supported

REM RICVIew
:rv
call skit_get_version rv 2>NUL
if NOT ERRORLEVEL = 1 goto done
echo RICView Not-supported

:done
