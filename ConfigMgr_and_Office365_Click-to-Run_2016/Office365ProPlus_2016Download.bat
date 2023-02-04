::From kevinisms.fason.org
@ECHO OFF
:: This Script will download the Office365ProPlus 2016 Install

:: Setup Shop
SET LogLocation=%TEMP%
SET XMLLOC=\\server\path\to\OfficeClick2Run\Office2016
SET SMTPTO=helpdesk@my.company.com
SET SMTPFROM=O365ProPlusC2RDownload@my.company.com
SET SMTPSMARTHOST=smtpsmarthost.my.company.com
SET VERSIONINFO=%TEMP%\O3652016Version.txt

::Cleanup from last run
DEL %LOGLOCATION%\%COMPUTERNAME%-20*.LOG
DEL %VERSIONINFO%

::Downloading
ECHO Downloading Office 365 2016
::  32-Bit
ECHO   Downloading 32-Bit
ECHO      Current Channel
ECHO ON
%~dp0setup2016.exe /download %XMLLOC%\download_X86_Current.xml
@ECHO      Deferred Channel
%~dp0setup2016.exe /download %XMLLOC%\download_X86_Deferred.xml
@ECHO      First Release for Current Channel 
%~dp0setup2016.exe /download %XMLLOC%\download_X86_First_Current.xml
@ECHO      First Release for Deferred Channel 
%~dp0setup2016.exe /download %XMLLOC%\download_X86_First_Deferred.xml
@ECHO OFF
ECHO.
::  64-Bit
ECHO   Downloading 64-Bit
ECHO      Current Channel 
ECHO ON
%~dp0setup2016.exe /download %XMLLOC%\download_X64_Current.xml
@ECHO      Deferred Channel 
%~dp0setup2016.exe /download %XMLLOC%\download_X64_Deferred.xml
@ECHO      First Release for Current Channel 
%~dp0setup2016.exe /download %XMLLOC%\download_X64_First_Current.xml
@ECHO      First Release for Deferred Channel 
%~dp0setup2016.exe /download %XMLLOC%\download_X64_First_Deferred.xml
@ECHO OFF
ECHO.

:: Get version info for email
ECHO( > %VERSIONINFO%
ECHO Office365ProPlus 2016 Click-to-Run Download Complete. See you next time! >> %VERSIONINFO%
ECHO( >> %VERSIONINFO%
ECHO Office365 2016 Versions at %XMLLOC% >> %VERSIONINFO%
ECHO( >> %VERSIONINFO%
ECHO Deferred Channel >> %VERSIONINFO%
DIR /A:D /B /O:N "%XMLLOC%\Deferred\Office\Data" >> %VERSIONINFO%
ECHO( >> %VERSIONINFO%
ECHO First Release for Deferred Channel >> %VERSIONINFO%
DIR /A:D /B /O:N "%XMLLOC%\FirstReleaseforDeferred\Office\Data" >> %VERSIONINFO%
ECHO( >> %VERSIONINFO%
ECHO Current Channel >> %VERSIONINFO%
DIR /A:D /B /O:N "%XMLLOC%\Current\Office\Data" >> %VERSIONINFO%
ECHO( >> %VERSIONINFO%
ECHO First Release for Current Channel >> %VERSIONINFO%
DIR /A:D /B /O:N "%XMLLOC%\FirstReleaseforCurrent\Office\Data" >> %VERSIONINFO%
ECHO( >> %VERSIONINFO%

::E-Mail results to GDM
ECHO.
ECHO Emailing results
"%~dp0blat.exe" %VERSIONINFO%  -to %SMTPTO% -f %SMTPFROM% -s "O365ProPlus 2016 C2R Download Results for %COMPUTERNAME%" -server %SMTPSMARTHOST%

