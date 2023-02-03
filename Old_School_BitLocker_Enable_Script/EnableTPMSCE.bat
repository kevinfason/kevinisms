@ECHO OFF
:: This BAT will call the CCTK files to change TPM settings to enabled for BITLOCKER support
:: By Kevin Fason
:: http://kevinisms.fason.org

::Sanity Check
::Check if script already executed. If so exit out
IF EXIST %WINDIR%\Temp\EnableTPMSCERan.txt GOTO EXIT
::This is the first run, creating dropper file for future executions to detect
TYPE NUL > %WINDIR%\Temp\EnableTPMSCERan.txt

::Busy Work 
:: Check OS Architecture and run correct version of downgrade app  
:CheckOS Arch  
 IF EXIST "%PROGRAMFILES(X86)%" (GOTO 64BIT) ELSE (GOTO 32BIT)  
             
:64BIT
::Set Password. This is needed for some settings such as TPM
%~dp0multiplatform_SetPassword_x64.exe
:: Set TPM settings
%~dp0multiplatform_Enable_TPM_x64.exe
:: Remove Password as its not needed anymore
%~dp0multiplatform_Reset_password_x64.exe
:: Set restart
%WINDIR%\System32\shutdown.exe /r /t 43200  /c "The IT Department has made hardware changes which requires a restart. Questions? Please open a ticket with NCC IT Support."
GOTO EXIT
 
:32BIT
::Set Password. This is needed for some settings such as TPM
%~dp0multiplatform_SetPassword.exe
:: Set TPM settings
%~dp0multiplatform_Enable_TPM.exe
:: Remove Password as its not needed anymore
%~dp0multiplatform_Reset_password.exe
:: Set restart
%WINDIR%\System32\shutdown.exe /r /t 43200  /c "The IT Department has made hardware changes which requires a restart. Questions? Please open a ticket with NCC IT Support."
GOTO EXIT

:EXIT