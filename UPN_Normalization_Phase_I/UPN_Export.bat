:: From HTTP://kevinisms.fason.org
:: This batch file will pull all user objects 
:: You will need adfind.exe in the same directory
:: Created by Kevin Fason
::
:: v1.0    01.23.2007
::         Initial Release

@ECHO OFF
CLS
IF "%1" == "/?" GOTO SYNTAX
IF "%1" == "" ECHO ERROR:  Domain Controller not entered
IF "%1" == "" GOTO SYNTAX



adfind -csv -t 500 -h %1 -b "dc=amr,dc=ch2m,dc=com" -f "&((objectcategory=person)(objectclass=user))" -nodn samaccountname userPrincipalName employeeID sn givenname name >> prod_List.CSV
ECHO Exported to prod_List.CSV
GOTO END

:SYNTAX
ECHO.
ECHO Syntax: mmadfind.BAT Domain_Controller
ECHO.
ECHO.
:END
