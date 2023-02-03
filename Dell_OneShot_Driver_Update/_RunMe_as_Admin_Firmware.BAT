@ECHO OFF
:: Wrapper for Dell Command Update CLI EXE
::  Written by Kevin Fason
::  v1.0 - Initial Release
:: http://kevinisms.fason.org

:: Setup Shop
SET LOGLOCATION=\\server\path\to\logs
CLS
ECHO.
ECHO Updating UEFI (BIOS) Firmware
ECHO.

::Removing local logs if found
ECHO - Removing old Log files if present
IF EXIST %PROGRAMDATA%\dell\CommandUpdate\ActivityLog.xml DEL %PROGRAMDATA%\dell\CommandUpdate\ActivityLog.xml
IF EXIST %PROGRAMDATA%\dell\CommandUpdate\inventory.xml DEL %PROGRAMDATA%\dell\CommandUpdate\inventory.xml
ECHO.

:: Since a BIOS update or controller update can trip BitLocker we will disable.
ECHO - Disabling Bitlocker if present
%WINDIR%\system32\manage-bde.exe -protectors -disable %SYSTEMDRIVE% > NUL
ECHO.

::Performin One Shot Driver update
:: Can append /silent to supress output
ECHO - Starting Dell Command Update CLI
ECHO.
"%~dp0dcu-cli.exe" /policy "%~dp0FirmwareOnly.xml"
SET DCUERRORLEVEL=%ERRORLEVEL%
ECHO.

::Disabled for production
::Copying logs initially to SERVER
REM ::ECHO Copying Log Files to %LOGLOCATION%
REM ::COPY /Y %PROGRAMDATA%\dell\CommandUpdate\ActivityLog.xml %LOGLOCATION%\%COMPUTERNAME%_ActivityLog.xml > NUL
REM ::COPY /Y %PROGRAMDATA%\dell\CommandUpdate\inventory.xml %LOGLOCATION%\%COMPUTERNAME%_inventory.xml > NUL
REM ::ECHO.

::Done. Show results based on what the script did.
ECHO Update complete
IF %DCUERRORLEVEL% == 0 ECHO Resuming Bitlocker if present
IF %DCUERRORLEVEL% == 0 %WINDIR%\system32\manage-bde.exe -enable %SYSTEMDRIVE% > NUL
IF %DCUERRORLEVEL% == 0 ECHO.
IF %DCUERRORLEVEL% == 1 ECHO Restart PC and then Resume Bitlocker and
IF %DCUERRORLEVEL% == 2 ECHO Fatal Error. Look up for more info and
IF %DCUERRORLEVEL% == 3 ECHO Invalid System. Look up for more info and
PAUSE

