@ECHO OFF
:: This Script will download the Office365ProPlus 2013 Install

:: Setup Shop
SET LogLocation=%TEMP%
SET XMLLOC=\\server\path\to\OfficeClick2Run\Office2013
SET X86XML=%XMLLOC%\download_X86.xml
SET X64XML=%XMLLOC%\download_X64.xml
SET SMTPTO=helpdesk@my.company.com
SET SMTPFROM=O365ProPlusC2RDownload@my.company.com
SET SMTPSMARTHOST=smtpsmarthost.my.company.com
SET VERSIONINFO=%TEMP%\O3652016Version.txt

::Cleanup from last run
DEL %LOGLOCATION%\%COMPUTERNAME%-20*.LOG
DEL %VERSIONINFO%

::Downloading
::  32-Bit
ECHO   Downloading 32-Bit 
%~dp0setup2013.exe /download %X86XML%

::  64-Bit
ECHO   Downloading 64-Bit
%~dp0setup2013.exe /download %X64XML%

:: Get version info for email
ECHO( > %VERSIONINFO%
ECHO Office365ProPlus 2013 Click-to-Run Download Complete. See you next time! >> %VERSIONINFO%
ECHO( >> %VERSIONINFO%
ECHO Office365 2013 Versions at %XMLLOC%\ >> %VERSIONINFO%
ECHO( >> %VERSIONINFO%
DIR /A:D /B /O:N "%XMLLOC%\Office\Data" >> %VERSIONINFO%
ECHO( >> %VERSIONINFO%


::E-Mail results to GDM
ECHO.
ECHO Emailing results
"%~dp0blat.exe" %VERSIONINFO%  -to %SMTPTO% -f %SMTPFROM% -s "O365ProPlus 2013 C2R Download Results for %COMPUTERNAME%" -server %SMTPSMARTHOST%

