@ECHO OFF
:: From http://kevinisms.fason.org
:: Downgrade Office 365 2013
:: by Kevin Fason
::  v1.0 - 11.23.2015 - Initial Release

CLS
IF "%1" == "/?" GOTO SYNTAX
IF "%1" == "" ECHO ERROR:  Downgrade Version not stated
IF "%1" == "" GOTO SYNTAX

:: Check OS Architecture and run correct version of downgrade app
:CheckOS Arch
IF EXIST "%PROGRAMFILES(X86)%" (GOTO 64BIT) ELSE (GOTO 32BIT)

:64BIT
%WINDIR%\sysnative\reg.exe add "HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\Configuration" /v "UpdateUrl" /t REG_SZ /d "\\my.company.com\path\to\files" /f
"%ProgramW6432%\Microsoft Office 15\ClientX64\officec2rclient.exe" /update user updatetoversion=%1
GOTO END

:32BIT
%WINDIR%\system32\reg.exe add "HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\Configuration" /v "UpdateUrl" /t REG_SZ /d ""\\my.company.com\path\to\files" /f
"%programfiles%\Microsoft Office 15\ClientX86\officec2rclient.exe" /update user updatetoversion=%1

GOTO END

:SYNTAX
ECHO.
ECHO Syntax: DowngradeOffice365_2013.bat 15.0.xxxx.yyyy
ECHO    Sample: DowngradeOffice365_2013.bat 15.0.4765.1003
ECHO.
ECHO.

:END