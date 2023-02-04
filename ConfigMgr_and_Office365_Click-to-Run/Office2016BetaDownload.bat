:: From Kevinisms.fason.org
@ECHO OFF
:: This Script will download the Office365ProPlus Install

:: Setup Shop
SET LogLocation=%TEMP%
SET X86XML=\\path\to\my\server\OfficeClick2Run\Office2016Beta\download_X86.xml
SET X64XML=\\path\to\my\server\OfficeClick2Run\Office2016Beta\download_X64.xml
SET SMTPTO=helpdesk@my.company.com
SET SMTPFROM=O365ProPlusC2RDownload@my.company.com
SET SMTPSMARTHOST=smtpsmarthost.my.company.com

::Cleanup from last run
DEL %LOGLOCATION%\%COMPUTERNAME%-20*.LOG

::Downloading
::  32-Bit
ECHO Downloading 32-Bit 
%~dp0setup2016.exe /download %X86XML%

::64-Bit
ECHO Downloading 64-Bit
%~dp0setup2016.exe /download %X64XML%

::E-Mail results to GDM
ECHO.
ECHO Emailing results
"%~dp0blat.exe" -body "Office365ProPlus 2016 Beta Click-to-Run Download Results attached. Please review. see you next time!" -to %SMTPTO% -f %SMTPFROM% -s "O365ProPlus C2R Download Results for %COMPUTERNAME%" -server %SMTPSMARTHOST% -attacht %LOGLOCATION%\%COMPUTERNAME%*.LOG