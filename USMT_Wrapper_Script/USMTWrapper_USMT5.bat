:: By http://kevinisms.fason.org
:: 	This batch file will run USMT 5 to capture user profiles.
::   		NOTE: You can precreate the USMTStore and/or USMTLogPath Variables to redirect this script
:: 		Created by Kevin Fason and Scott Freeman
::
:: 	v2.0    03.29.2013
::         	This version uses MIGDocs.xml to gather user data versus MigUser.xml
::		XML Files Used: MigDocs, MigApp, IncludeExclude.xml	
::

@ECHO OFF
CLS

::Get Correct Processor Architecture from Registry since %Processor_Architecture% can get the wrong arch
::When run via Package Downloader...
@echo off
setlocal ENABLEEXTENSIONS
set KEY_NAME=HKLM\System\CurrentControlSet\Control\Session Manager\Environment
set VALUE_NAME=PROCESSOR_ARCHITECTURE
for /F "usebackq tokens=3" %%A IN (`reg query "%KEY_NAME%" /v "%VALUE_NAME%" 2^>nul ^| find "%VALUE_NAME%"`) do (
  SET Proc_Arch=%%A
)

:: Get Script Start time for later use
FOR /F "tokens=*" %%i in ('TIME /T') do SET USMTScriptStartTime=%%i

::Setup Shop
:: Copy the USMT Package locally as we hit the 256 character command line wall otherwise
ECHO.
ECHO Copying USMT(5) Files Local to Windows\Temp and beginning. Please wait...
mkdir %WINDIR%\TEMP\USMT /f >nul 2>&1

"%WINDIR%\System32\xcopy.exe" "%~dp0*.*" "%WINDIR%\TEMP\USMT" /I /E /Y /Q > NUL 
%SYSTEMDRIVE%
CD %WINDIR%\TEMP\USMT

:: Check for XP and if found copy the USMT manifest files to System32
:: Printer Migration can fail if migrating from XP > newer because XP doesnt have the USMT manifest files
:: 

Echo.
Echo Checking for XP

ver | find "5.1" > nul
if %ERRORLEVEL% == 0 echo XP found, copying manifest files to system32
if %ERRORLEVEL% == 0 "%WINDIR%\System32\xcopy.exe" "%WINDIR%\TEMP\USMT\%Proc_Arch%\DLManifests" "%WINDIR%\System32\DLManifests" /I /E /Y /Q >NUL 
if %ERRORLEVEL% == 1 Echo XP not found 


Echo.
Echo ...Ready to begin
Echo.


::Setup Universal Variables
SET USMTProgramPath=%Proc_Arch%
SET USMTMigFiles=/i:%USMTProgramPath%\migdocs.xml /i:%USMTProgramPath%\migapp.xml /i:%USMTProgramPath%\IncludeExclude.xml


::Determine Which section to run
::  LoadState restores data
::  ScanState collects data

:USMTFoundPrompt
IF EXIST %SYSTEMDRIVE%\USMTStore set /p answer='C:\USMTStore' folder detected, Is this the new computer. (Y/N)?
Echo.
IF NOT EXIST %SYSTEMDRIVE%\USMTStore GOTO SCANSTATE
If /i "%answer:~,1%" EQU "Y" GOTO LOADSTATE
If /i "%answer:~,1%" EQU "N" RD %SYSTEMDRIVE%\USMTStore
If /i "%answer:~,1%" EQU "" GOTO USMTFoundPrompt

	
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
IF "%1" == "" ECHO What is the new Computername?
IF "%1" == "" ECHO.
IF "%1" == "" ECHO    Sample: COMPUTERNAME
IF "%1" == "" ECHO    Sample: 10.10.10.12
IF "%1" == "" ECHO.
IF "%1" == "" set /p USMTNEWMACHINE= >%USMTProgramPath%\Input1.bat
::IF "%1" == "" %USMTProgramPath%\sed -e "s/^/SET USMTNEWMACHINE=/" -e "q" >%USMTProgramPath%\Input1.bat
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
%USMTProgramPath%\scanstate.exe %USMTStore%\ %USMTMigFiles% /c /v:5 /l:%USMTLogPath%\ScanState_%COMPUTERNAME%.log /progress:%USMTLogPath%\ScanStateProgress_%COMPUTERNAME%.log /uel:30 /ue:"%COMPUTERNAME%\*" /efs:copyraw /vsc /o 
         
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
%USMTProgramPath%\loadstate.exe %USMTStore%\ %USMTMigFiles% /c /v:5 /l:%USMTLogPath%\LoadState_%COMPUTERNAME%.log /progress:%USMTLogPath%\LoadStateProgress_%COMPUTERNAME%.log  /uel:30 /ue:"%COMPUTERNAME%\*"

::TODO - Errorlevel handling? 

::Copy Logs to Store Path
IF %Proc_Arch%==x86 "%WINDIR%\System32\xcopy.exe" "%USMTStore%\*.log" "%WINDIR%\System32\CCM\Logs" /I /E /Y  /Q > NUL
IF %Proc_Arch%==AMD64 "%WINDIR%\System32\xcopy.exe" "%USMTStore%\*.log" "%WINDIR%\SysWOW64\CCM\Logs" /I /E /Y /Q > NUL

                                                                        
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

set timestamp=%TIME::=%
set timestamp=%timestamp: =%
set timestamp=%date:~-4,4%%date:~-7,2%%date:~-10,2%_%timestamp:~0,-3%

ren %SYSTEMDRIVE%\USMTstore "USMTStore_%timestamp%"
IF EXIST %SYSTEMDRIVE%\USMTStore ren %SYSTEMDRIVE%\USMTStore "USMTStore_OLD"

Echo.
ECHO The C:\USMTStore is being renamed so this system can be migrated in the future
Echo.
ECHO The new name is: %SYSTEMDRIVE%\USMTStore_%timestamp% or %SYSTEMDRIVE%\USMTStore_OLD
Echo.
ECHO If you must run USMT (Loadstate) again, rename back to USMTStore
Echo.
ECHO      Please delete this folder when no longer needed.

GOTO END

:END
Echo.
PAUSE

