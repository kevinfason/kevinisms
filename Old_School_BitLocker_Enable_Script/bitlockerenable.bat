@ECHO OFF
:: This BAT will enable Bitlocker on all drives found. You must set GPO policy so it uses your
::  moethod of choice.
:: By Kevin Fason
:: http://kevinisms.fason.org

::look for drives
setlocal ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION
    for /f "delims=" %%i in ('^
        echo list volume ^|^
        diskpart ^|^
        findstr Volume ^|^
        findstr /v ^
        /c:"Volume ###  Ltr  Label        Fs     Type        Size     Status     Info"^
       ') do (
        set "line=%%i"
        set letter=!line:~15,1!
        set fs=!line:~32,7!
        if not "       "=="!fs!" (
            if not " "=="!letter!" (
                call :Encrypt !letter!
            )
        )
    )
GOTO EXIT

:ENCRYPT
SET DRIVELETTER=%1:

::Detecting if Bitlocker is already on
%WINDIR%\System32\manage-bde.exe -status %1 | FIND "Protection On" > nul2
IF "%ERRORLEVEL%"=="0" EXIT/B

ECHO.
ECHO Encrypting Volume %DRIVELETTER% your PC, be patient . . .
ECHO.
ECHO  There is no Need to write down the numerical password below
ECHO.
TITLE Encrypting your PC, be patient . . .

::Create Recovery Key
ECHO Create Recovery Key
%WINDIR%\System32\manage-bde.exe -protectors -add %DRIVELETTER% -recoverypassword

::Create TPM Key
ECHO Create TPM Key
%WINDIR%\System32\manage-bde.exe -protectors -add %DRIVELETTER% -tpm

::Enable Bitlocker on Windows Drive
ECHO Enable Bitlocker on Windows Drive
%WINDIR%\System32\manage-bde.exe -on %DRIVELETTER%

Set BLEnabled=YES
EXIT /B

:EXIT
IF %BLENABLED%==YES %WINDIR%\System32\shutdown.exe /r /t 300  /c "IT Department made a change and your workstation will restart in 5 Mins. Questions? Please open a ticket with NCC IT Support."

