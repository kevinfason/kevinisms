:: by http://kevinisms.fason.org
:: This batch file will run USMT to capture user profiles.
::   NOTE: You can precreate the USMTStore and/or USMTLogPath Variables to redirect this script
:: Created by Kevin Fason and Scott Freeman
::
:: v1.0    09.29.2011
::         Initial Release
@ECHO OFF
CLS

:: Get Script Start time for later use
FOR /F "tokens=*" %%i in ('TIME /T') do SET USMTScriptStartTime=%%i

::Setup Shop
:: Copy the USMT Package locally as we hit the 256 character command line wall otherwise
ECHO.
ECHO Copying USMT Files Local to local TEMP and running DOS version. Please wait...
mkdir %WINDIR%\TEMP\USMT /f >nul 2>&1

"%WINDIR%\System32\xcopy.exe" "%~dp0*.*" "%WINDIR%\TEMP\USMT" /I /E /Y /Q > NUL 
%SYSTEMDRIVE%
CD %WINDIR%\TEMP\USMT
Echo ...Copy complete

taskkill /im USMTsetup.exe /f >nul 2>&1


::Setup Universal Variables
SET USMTProgramPath=%PROCESSOR_ARCHITECTURE%
SET USMTMigFiles=/i:%USMTProgramPath%\miguser.xml /i:%USMTProgramPath%\migapp.xml /i:%USMTProgramPath%\wallpaper.xml /i:%USMTProgramPath%\MMSettings.xml /config:%USMTProgramPath%\config.xml


::Determine Which section to run
::  LoadState restores data
::  ScanState collects data
IF EXIST %SYSTEMDRIVE%\USMTStore GOTO LoadState







:ScanState

::Sanity Checks
REM IF "%1" == "" ECHO "ERROR:  New Computer Not entered"
REM IF "%1" == "" GOTO SYNTAX
REM IF "%1" == "/?" GOTO SYNTAX

::Copy the CMD Line variable (if present) into our variable being used
SET USMTNEWMACHINE=%1

:: Prompt for new computer if nothing previuosly set.
::   TODO - More elegant way of doing this?
IF "%1" == "" ECHO.
IF "%1" == "" ECHO What is the new Computername or IP address?
IF "%1" == "" ECHO.
IF "%1" == "" ECHO    Sample: USMTWrapper.bat COMPUTERNAME
IF "%1" == "" ECHO    Sample: USMTWrapper.bat IPADDRESS
IF "%1" == "" ECHO.
IF "%1" == "" %USMTProgramPath%\sed -e "s/^/SET USMTNEWMACHINE=/" -e "q" >%USMTProgramPath%\Input1.bat
IF "%1" == "" CALL %USMTProgramPath%\Input1.bat

:: Setup Scanstate specific variables
:: Can preset the two following variables before running Script to override defaults
REM IF "%USMTStore%" == "" SET USMTStore=\\%1\C$\USMTStore
IF "%USMTStore%" == "" SET USMTStore=\\%USMTNEWMACHINE%\C$\USMTStore
REM SET USMTStore=\\%1\C$\USMTStore
REM SET USMTLogPath=%WINDIR%\TEMP\USMT
IF "%USMTLogPath%" == "" SET USMTLogPath=%USMTStore%




:: Run Scanstate
::     TODO - Insert status window?
::     TODO - Update Switch comments
:: Switches (http://technet.microsoft.com/en-us/library/dd560781%28WS.10%29.aspx)
::    /c - Continue on errors
::    /hardlinks - use hardlinks instead of creating a store
::    /nocompress - do not compress, must be used with hardlinks switch
::    /uel:30 - Get all users who have logged in within the last 30 days
::    /ue:ch2mhill\%USERNAME% - Excludes the logged in user
::      TODO - This switch will prevent the IT person form being included, however it means the assigned 
::             User cannot run this script!!
::    /ui:CH2MHILL\* - Get all CH2MHILL Domain Users
::      TODO - Logic for other domains from mergers etc
::    /v:5 - Verbosity level
::    /l - Log locations
::    /efs:copyraw - Copy any Encrypted Files found
::    /o - Overrights any data in the USMT store
::    /i - Include XML templates to capture or exclue specific data

:: Call the correct Arch Scanstate.
%USMTProgramPath%\scanstate.exe %USMTStore% %USMTMigFiles% /o /v:5 /l:%USMTLogPath%\ScanState_%COMPUTERNAME%.log /progress:%USMTLogPath%\ScanStateProgress_%COMPUTERNAME%.log /c /uel:30 /ue:"%COMPUTERNAME%/*" /ue:"CH2MHILL\%USERNAME%" /efs:copyraw /vsc
         
::Error handling
IF [NOT] %ERRORLEVEL% == 0 GOTO ERROR

::Copy Logs to Store Path
REM "%WINDIR%\System32\xcopy.exe" ".\*.log" "%USMTStore%" /I /E /Y 

::TODO - Delete %WINDIR%\TEMP\USMT?? Leave as a backup for logs etc?

:: Get Script Stop time
FOR /F "tokens=*" %%i in ('TIME /T') do SET USMTScriptStopTime=%%i

:ERROR
ECHO.
ECHO Error Detected. Please act on the above error or contact support.
GOTO END

:: Call USMTLogZip to Copy logs to network location
:: USMTLogZip %1,SCAN


::Closeup Shop
ECHO.
ECHO All Done!
ECHO If needed, the logs are located at %USMTStore% and called
ECHO ScanState_%COMPUTERNAME%.log
ECHO.
ECHO Run This script from the new computer to restore data
ECHO. 
ECHO USMT Wrapper Started at %USMTScriptStartTime%
ECHO.             Finished at %USMTScriptStopTime% 
ECHO.
GOTO END



:LoadState

ECHO ScanState already ran on this machine. Performing LoadState
ECHO.

::Setup LoadState Variables
IF "%USMTStore%" == "" SET USMTStore=%SYSTEMDRIVE%\USMTStore
SET USMTLogPath=%USMTStore%

:: Run LoadState
::     TODO - Insert status window?
::     TODO - Update Switch comments
:: Switches (http://technet.microsoft.com/en-us/library/cc766226%28WS.10%29.aspx)
::    /c - Continue on errors
::    /auto - autofind XML files
::    /l - Log locations

:: Call the correct Arch LoadState.
%USMTProgramPath%\loadstate.exe %USMTStore%\ /c /v:5 /l:%USMTLogPath%\LoadState_%COMPUTERNAME%.log /progress:%USMTLogPath%\LoadStateProgress_%COMPUTERNAME%.log %USMTMigFiles% /uel:30 /ue:* /ue:"CH2MHILL\%USERNAME%" /lac

::TODO - Errorlevel handling? 

::Copy Logs to Store Path
IF %PROCESSOR_ARCHITECTURE%==x86 "%WINDIR%\System32\xcopy.exe" "%USMTStore%\*.log" "%WINDIR%\System32\CCM\Logs" /I /E /Y  /Q > NUL
IF %PROCESSOR_ARCHITECTURE%==AMD64 "%WINDIR%\System32\xcopy.exe" "%USMTStore%\*.log" "%WINDIR%\SysWOW64\CCM\Logs" /I /E /Y /Q > NUL

:: Call USMTLogZip to Copy logs to network location
:: call USMTLogZip.exe NoPC,LOAD 
                                                                        
::Delete Store
::  TODO - simple del c:\USMTStore or use the util in USMT?
                 
:: Get Script Stop time
FOR /F "tokens=*" %%i in ('TIME /T') do SET USMTScriptStopTime=%%i
   
::Close Shop
ECHO.
ECHO All data transfered from the old machine 
ECHO If needed, the logs are located in the SCCM Logs folder and called
ECHO LoadState_NEWMACHINENAME.log and ScanState_OLDMACHINENAME.LOG
ECHO.
ECHO USMT Wrapper Started at %USMTScriptStartTime%
ECHO             Finished at %USMTScriptStopTime% 
ECHO.
ECHO Restart computer for all changes to be applied.
ECHO.


:: Rename Local USMTStore folder
ren c:\USMTstore USMtStoreOld


GOTO END

:END
PAUSE

