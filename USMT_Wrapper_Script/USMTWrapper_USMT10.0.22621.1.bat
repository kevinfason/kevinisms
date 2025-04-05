@ECHO OFF
CLS
:: By Kevin Fason
::     v1.0 08.01.2024 
::          Initial Release
::          Adopted from my old wrapper at http://kevinisms.fason.org

::Setup Shop
:: Detecting Admin rights
echo Administrative permissions required. Detecting permissions...
 net file >nul 2>&1
 if [%errorLevel%] == [0] (
 	echo Success: Administrative permissions confirmed.
 ) else (
 	echo Failure: Current permissions inadequate.
    goto NOADMINDETECTED
 )

:: Copy the USMT Package locally for performance
ECHO.
ECHO Copying USMT Files Locally to Windows\Temp and beginning. Please wait...
mkdir %WINDIR%\TEMP\USMT /f >nul 2>&1
"%WINDIR%\System32\xcopy.exe" "%~dp0*.*" "%WINDIR%\TEMP\USMT" /I /E /Y /Q > NUL 
%SYSTEMDRIVE%
CD %WINDIR%\TEMP\USMT

Echo.
Echo ...Ready to begin
Echo.

::Setup Universal Variables
:: Hardcoding to AMD64 as we should have zero X86 or ARM64 instances. There is no X86 in W11
:: Leaving framework in should there be any ARM64 in the future
:: SET USMTProgramPath=%PROCESSOR_ARCHITECTURE%
SET USMTProgramPath=AMD64
:: Declare relevent XML files
::  Most are sourced from EhlerTech with some defaults. Will create COMPANY specific one if needed
::  Use MigDocs OR MigUser, not both
::  MIG detail can be found here https://learn.microsoft.com/en-us/windows/deployment/usmt/understanding-migration-xml-files
SET USMTMigFiles=/i:%USMTProgramPath%\MigDocs.xml /i:%USMTProgramPath%\MigApp.xml /i:%USMTProgramPath%\ExcludeDrives_D_to_Z.xml /i:%USMTProgramPath%\Xclude_Network_Printers.xml /i:%USMTProgramPath%\Xclude_Windows_Defender.xml /i:%USMTProgramPath%\ExcludeSystemFolders.xml
:: Add IT Support staff to exclusions. If an IT Support person is migrating they will need to be called out as as a switch. 
::  IT Staff should have elevated accounts so you can pass 'samaccountname*' to get all of this users accounts. Or do two captures.
::  Hopefully IT staff escelated accounts are in standard syntax such as 'DOMAIN\samaccountname_admin' so can be ignored by default easily
:: Could add support to look for this use case however it would have to be ran by different tech so all accounts are unmounted.
:: Note there is a scanstate switch to skip any local accounts and executing account below
:: TODO do AD group extract and compare to local user repo to dynamically add?
SET USMTSCANUE=/ue:"DOMAIN\samaccountname1*" /ue:"DOMAIN\samaccountname2*" /ue:"DOMAIN\samaccountname3*" /ue:"DOMAIN\*admin"

::Determine Which section to run
::  LoadState restores data
::  ScanState collects data

:USMTFoundPrompt
:: Looking for an existing USMTStore folder locally
:: Only looking at first few volumes to prevent parsing any mapped UNC paths
if defined USMTStorePath (
    echo USMTStore folder path is already set to %USMTStorePath%.
) else (
    :: Search for USMTStore across drives C through G
    for %%a in (C D E F G) do (
        if exist "%%a:\USMTStore\" (
            set "USMTStorePath=%%a:\USMTStore"
            REM goto LoadState
        )
    )
)

:: Prompt if USMTStore was not found on any drive
echo.
if defined USMTStorePath (
    set /p answer="'%USMTStorePath%' folder detected, Is this the new computer? (Y/N)? "
)       
if /i "%answer:~0,1%" == "Y" ( 
        GOTO LOADSTATE
    )
if /i "%answer:~0,1%" == "N" (
        echo renaming %USMTStorePath% to %USMTStorePath%_old and starting capture
        REMrd /s /q "%USMTStorePath%"
        goto scanstate
    ) 
	
:ScanState
TITLE Performing USMT Capture
::Sanity Checks
:: Prompt for new computer if nothing previously set.
if "%~1"=="" (
    echo.
    echo What is the USMT Store path?
    echo.
    echo    Sample: D:\
    echo    Sample: \\server\path\
    echo    Sample: \\newcomputer\C$\ 
    set /p USMTStorePath="Enter the USMT Store path: "
    ) else (
    set USMTStorePath=%~1
)
:: Cannot do this within the above if then
set USMTStorePath=%USMTStorePath%USMTStore
echo Selected USMT Store Path: %USMTStorePath%

:: Setup Scanstate specific variables
:: Determine if we are a W10 or W11 source instance
for /f "tokens=6 delims=. " %%a in ('ver') do if %%a GEQ 22000 (SET W10orW11=/i:%USMTProgramPath%\win11.xml) ELSE IF %%a LEQ 19999 (SET W10orW11=/i:%USMTProgramPath%\win10.xml)
:: TODO Add OneDrive Detection to pass /i:%USMTProgramPath%\ExcludeOneDriveRedirFolders.xml so we ignore those files
:: Setting Logpath if not declared earlier
SET USMTLogPath=%USMTStorePath%

:: Determine Userinclude. 
:: If samaccountname is passed to script we only export that user. 
::  Otherwise we export any Domain users that logged in to this instance in the last 30 days
if "%~2"=="" (
    echo.
    echo No sAMAccountname identified
    echo Scanning all users who logged in within the last 30 days
    set USMTUser=/uel:30
) else (
    echo.
    echo The sAMAccountname %~2 was passed to the script. Only Capturing DOMAIN\%~2
    echo.
    set USMTUser=/ue:*\* /ui:DOMAIN\%~2
)

:: Run Scanstate
::     TODO - Insert status window?
::     TODO - Update Switch comments
:: Switches (http://technet.microsoft.com/en-us/library/dd560781%28WS.10%29.aspx)
::    /c - Continue on errors
::    /hardlinks - use hardlinks instead of creating a store
::    /nocompress - do not compress, must be used with hardlinks switch
::    /uel:30 - Get all users who have logged in within the last 30 days
::    /ue:DOMAIN\%USERNAME% - Excludes the logged in user
::    /ui:DOMAIN\* - Get all Domain Users
::    /ui:DOMAIN\sAMAccountname - Get only this Domain user
::    /v:5 - Verbosity level
::    /l - Log locations
::    /efs:copyraw - Copy any Encrypted Files found
::    /o - Overwrites any data in the USMT store
::    /i - Include XML templates to capture or exclue specific data

:: Get Section Start time for later use
FOR /F "tokens=*" %%i in ('TIME /T') do SET USMTScriptStartTime=%%i

:: Call the correct Arch Scanstate.
%USMTProgramPath%\scanstate.exe %USMTStorePath%\ %USMTMigFiles% %W10orW11% %USMTSCANUE% %USMTUser% /ue:%COMPUTERNAME%\* /c /v:5 /l:%USMTLogPath%\ScanState_%COMPUTERNAME%.log /progress:%USMTLogPath%\ScanStateProgress_%COMPUTERNAME%.log  /efs:copyraw /vsc /o 
set ScanStateErrorLevel=%ERRORLEVEL%

::Copy Logs to Store Path
REM "%WINDIR%\System32\xcopy.exe" ".\*.log" "%USMTStorePath%" /I /E /Y 

::TODO - Delete %WINDIR%\TEMP\USMT?? Leave as a backup for logs etc?

:: Get Script Stop time
FOR /F "tokens=*" %%i in ('TIME /T') do SET USMTScriptStopTime=%%i


::Closeup Shop
ECHO.
ECHO All Done!
echo.
ECHO If needed, the logs are located at %USMTStorePath% and called
ECHO ScanState_%COMPUTERNAME%.log
ECHO.
ECHO Run This script from the new computer to restore data
ECHO. 
ECHO USMT Wrapper Started at %USMTScriptStartTime%
ECHO.             Finished at %USMTScriptStopTime% 
ECHO.
IF %ScanStateErrorLevel% NEQ 0 goto Error
GOTO END


:LoadState
TITLE Performing USMT restore
ECHO ScanState already ran for this machine. Performing LoadState
ECHO.
Echo Selected USMT Store Path: %USMTStorePath%
Echo.

::Setup LoadState Variables
IF "%USMTStorePath%" == "" SET USMTStore=%SYSTEMDRIVE%\USMTStore
SET USMTLogPath=%USMTStorePath%

:: Run LoadState
::     TODO - Insert status window?
::     TODO - Update Switch comments
:: Switches (http://technet.microsoft.com/en-us/library/cc766226%28WS.10%29.aspx)
::    /c - Continue on errors
::    /l - Log locations
::    /v:5 - Verbosity level

:: Get Section Start time for later use
FOR /F "tokens=*" %%i in ('TIME /T') do SET USMTScriptStartTime=%%i

:: Call the correct Arch LoadState.             
%USMTProgramPath%\loadstate.exe %USMTStorePath% %USMTMigFiles% %USMTSCANUE% /uel:30 /c /v:5 /l:%USMTLogPath%\LoadState_%COMPUTERNAME%.log /progress:%USMTLogPath%\LoadStateProgress_%COMPUTERNAME%.log
set LoadStateErrorLevel=%ERRORLEVEL%

::Copy Logs to ConfigMgr Logs folder
"%WINDIR%\System32\xcopy.exe" "%USMTStorePath%\*.log" "%WINDIR%\System32\CCM\Logs" /I /E /Y  /Q > NUL

                                                                        
::Delete Store
::  TODO - simple del instead of rename? Or use the util in USMT?
                 
:: Get Script Stop time
FOR /F "tokens=*" %%i in ('TIME /T') do SET USMTScriptStopTime=%%i
   
::Closeup Shop
ECHO.
ECHO All data transfered from the old machine 
ECHO If needed, the logs are located in the ConfigMgr Logs folder and called
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


move "%USMTStorePath%" "%USMTStorePath%_%timestamp%" > NUL

Echo.
ECHO The USMTStore is being renamed so this system can be migrated in the future
Echo.
ECHO The new name is: %USMTStorePath%_%timestamp% 
Echo.
ECHO If you must run USMT (Loadstate) again, rename back to %USMTSTOREPath%
Echo.
ECHO      Please delete this folder when no longer needed.
IF %LoadStateErrorLevel% NEQ 0 goto Error
GOTO END

:NOADMINDETECTED
 COLOR 04
 ECHO.  
 ECHO ####### WARNING: ADMINISTRATOR PRIVILEGES REQUIRED #########  
 ECHO This script must be run as a member of local administrator 
 ECHO to work properly!  
 ECHO If you're seeing this then right click on the shortcut 
 ECHO and select "Run As Administrator".
 ECHO Alternatively (while holding shift) right click 
 ECHO %WINDIR%\System32\CMD.EXE 
 ECHO then select "Run As Another User" and enter a privileged account
 ECHO and run the installer again.
 ECHO ##########################################################  
 ECHO.
 PAUSE  
 COLOR 
 EXIT /B 666 

:ERROR
COLOR 04
ECHO.
ECHO Error Detected. Please act on the above error or contact support. 
ECHO More detailed logs may be in this folder
GOTO END

:END
Echo.
PAUSE
COLOR


