:: by http://kevinisms.fason.org

@ECHO OFF
:: Script to reset WUA on a client. It will revert it to a fresh state.
:: It will stop all services, backup relevent data. setup wsus and start all services back up
::  Written by Kevin Fason
:: version 1.0
::    12/14/2011
:: version 1.1
::    10/14/2012
::    - Added backup of Reg file
::    - Copy Distro folder instead of deleting.


ECHO.
ECHO This Script will reset WSUS Client . Please Wait...

ECHO Stopping relevent services
ECHO.
ECHO   Stopping Windows Update Agent Service
NET stop wuauserv

ECHO.
ECHO   Stopping BITS Service
NET stop bits

:: Un Comment the next 3 lines if you use ConfigMgr
:: ECHO.
:: ECHO   Stopping SCCM Host Agent Service
:: NET stop ccmexec

ECHO.
ECHO Backing up SoftwareDistribution folder
MOVE %WINDIR%\SoftwareDistribution %WINDIR%\SoftwareDistribution.old

ECHO.
ECHO Deleting WindowsUpdate Log from %WINDIR%
DEL /F /Q %WINDIR%\WIndowsUpdate.LOG

ECHO.
ECHO Backing up WindowsUpdate Registry entires to %TEMP%
REG EXPORT "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate" %TEMP%\WSUSResetBackup.REG

ECHO.
ECHO Removing WSUS Client ID from Registry
ECHO.
ECHO    Deleting AccountDomainSid if present
REG DELETE "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate" /v AccountDomainSid /f

ECHO    Deleting PingId if present
REG DELETE "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate" /v PingID /f

ECHO    Deleting SusClientId if present
REG DELETE "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate" /v SusClientId /f

ECHO    Deleting SusClientIdValidation if present
REG DELETE "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate" /v SusClientIdValidation /f

ECHO    Deleting LastWaitTimeout if present
REG DELETE "HKLM\Software\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update" /v LastWaitTimeout /f 

ECHO    Deleting DetectionStartTime if present
REG DELETE "HKLM\Software\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update" /v DetectionStartTime /f 

ECHO    Deleting NextDetectionTime if present
REG DELETE "HKLM\Software\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update" /v NextDetectionTime /f 

:: Remove this section if you do not use GRoup Policy for server settings
ECHO.
ECHO Force Group Policy Update For WSUS Server Settings
gpupdate /force

ECHO Starting relevent services
ECHO.

:: Un Comment the next 3 lines if you use ConfigMgr
:: ECHO.
:: ECHO    Starting SCCM Host Agent Service
:: NET start ccmexec

ECHO.
ECHO    Starting BITS Service
NET start bits

ECHO.
ECHO    Starting Windows Update Service
NET start wuauserv

ECHO.
ECHO Resetting Authorization Cookie and Invoking Update Detection
wuauclt /a /resetauthorization /detectnow
                                               
ECHO.
ECHO All Done
PAUSE
