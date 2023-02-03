:: https://kevinisms.fason.org
:: This will take an existing Windows 7 WIM and run against Jason Sandys
::  https://home.configmgrftw.com/building-a-windows-7-image-revisited/
:: process to create a Final supported Windows 7 version
:: You need to get the install.wim from a Windows 7 install media such as AIO.
:: Modify to your needs from your source install.wim by looking at its index

::Setup Shop
SET ORIGINALISO=_Windows_7_SP1_AIO
SET ORIGINALWIM=%ORIGINALISO%\Sources\Install.WIM

:: Export From AIO WIM

TITLE Exporting WIMs
dism /Export-Image /SourceImageFile:%ORIGINALWIM% /SourceIndex:1 /DestinationImageFile:image\STARTER.wim /DestinationName:"Windows 7 Starter" /Compress:max
dism /Export-Image /SourceImageFile:%ORIGINALWIM% /SourceIndex:2 /DestinationImageFile:image\HOMEBASIC.wim /DestinationName:"Windows 7 Home Basic" /Compress:max
dism /Export-Image /SourceImageFile:%ORIGINALWIM% /SourceIndex:3 /DestinationImageFile:image\HOMEPREMIUM.wim /DestinationName:"Windows 7 Home Premium" /Compress:max
dism /Export-Image /SourceImageFile:%ORIGINALWIM% /SourceIndex:4 /DestinationImageFile:image\PROFESSIONAL.wim /DestinationName:"Windows 7 Home Professional" /Compress:max
dism /Export-Image /SourceImageFile:%ORIGINALWIM% /SourceIndex:5 /DestinationImageFile:image\ULTIMATE.wim /DestinationName:"Windows 7 Home Ultimate" /Compress:max
dism /Export-Image /SourceImageFile:%ORIGINALWIM% /SourceIndex:6 /DestinationImageFile:image\HOMEBASIC_x64.wim /DestinationName:"Windows 7 Home Basic x64" /Compress:max
dism /Export-Image /SourceImageFile:%ORIGINALWIM% /SourceIndex:7 /DestinationImageFile:image\HOMEPREMIUM_x64.wim /DestinationName:"Windows 7 Home Premium x64" /Compress:max
dism /Export-Image /SourceImageFile:%ORIGINALWIM% /SourceIndex:8 /DestinationImageFile:image\PROFESSIONAL_x64.wim /DestinationName:"Windows 7 Professional x64" /Compress:max
dism /Export-Image /SourceImageFile:%ORIGINALWIM% /SourceIndex:9 /DestinationImageFile:image\ULTIMATE_x64.wim /DestinationName:"Windows 7 Ultimate x64" /Compress:max
dism /Export-Image /SourceImageFile:%ORIGINALWIM% /SourceIndex:10 /DestinationImageFile:image\ENTERPRISE.wim /DestinationName:"Windows 7 Enterprise" /Compress:max
dism /Export-Image /SourceImageFile:%ORIGINALWIM% /SourceIndex:11 /DestinationImageFile:image\ENTERPRISE_x64.wim /DestinationName:"Windows 7 Enterprise x64" /Compress:max

:: Copy Relevent WIMs for injection
TITLE Copying WIMs for injection
COPY image\PROFESSIONAL_x64.WIM image\Professional_Final_x64.WIM
COPY image\ENTERPRISE_x64.WIM image\Enterprise_Final_x64.WIM
COPY image\ULTIMATE_x64.WIM image\Ultimate_Final_x64.WIM
COPY image\HOMEPREMIUM_x64.WIM image\HomePremium_Final_x64.WIM

:: Calling Jasons Script against relevent WIMs
ECHO Calling Jasons script against relevent WIMs
ECHO It will pause at the beginning of each run
ECHO Unless you REM the PAUSE statement

TITLE Creating Final Professional
Win7Image-Update image\Professional_Final_x64.WIM .\Mount
TITLE Creating Final Enterprise
Win7Image-Update image\Enterprise_Final_x64.WIM .\Mount
TITLE Creating Final Ultimate
Win7Image-Update image\Ultimate_Final_x64.WIM .\Mount
TITLE Creating Final Home Premium
Win7Image-Update image\HomePremium_Final_x64.WIM .\Mount

:: Create new WIM with Final and most used at top
TITLE Creating Final WIM
dism /Export-Image /SourceImageFile:image\Professional_Final_x64.WIM /SourceIndex:1 /DestinationImageFile:image\installfinal.wim /DestinationName:"Windows 7 Professional Final 01/2020 x64" /Compress:max
dism /Export-Image /SourceImageFile:image\Enterprise_Final_x64.WIM /SourceIndex:1 /DestinationImageFile:image\installfinal.wim /DestinationName:"Windows 7 Enterprise Final 01/2020 x64" /Compress:max
dism /Export-Image /SourceImageFile:image\Ultimate_Final_x64.WIM /SourceIndex:1 /DestinationImageFile:image\installfinal.wim /DestinationName:"Windows 7 Ultimate Final 01/2020 x64" /Compress:max
dism /Export-Image /SourceImageFile:image\HomePremium_Final_x64.WIM /SourceIndex:1 /DestinationImageFile:image\installfinal.wim /DestinationName:"Windows 7 Home Premium Final 01/2020 x64" /Compress:max

:: Add stock SP1 WIMs back in
dism /Export-Image /SourceImageFile:image\PROFESSIONAL_x64.wim /SourceIndex:1 /DestinationImageFile:image\installfinal.wim /DestinationName:"Windows 7 Professional x64" /Compress:max
dism /Export-Image /SourceImageFile:image\ENTERPRISE_x64.wim /SourceIndex:1 /DestinationImageFile:image\installfinal.wim /DestinationName:"Windows 7 Enterprise x64" /Compress:max
dism /Export-Image /SourceImageFile:image\ULTIMATE_x64.wim /SourceIndex:1 /DestinationImageFile:image\installfinal.wim /DestinationName:"Windows 7 Ultimate x64" /Compress:max
dism /Export-Image /SourceImageFile:image\HOMEPREMIUM_x64.wim /SourceIndex:1 /DestinationImageFile:image\installfinal.wim /DestinationName:"Windows 7 Home Premium x64" /Compress:max
dism /Export-Image /SourceImageFile:image\HOMEBASIC_x64.wim /SourceIndex:1 /DestinationImageFile:image\installfinal.wim /DestinationName:"Windows 7 Home Basic x64" /Compress:max
dism /Export-Image /SourceImageFile:image\ENTERPRISE.wim /SourceIndex:1 /DestinationImageFile:image\installfinal.wim /DestinationName:"Windows 7 Enterprise" /Compress:max
dism /Export-Image /SourceImageFile:image\PROFESSIONAL.wim /SourceIndex:1 /DestinationImageFile:image\installfinal.wim /DestinationName:"Windows 7 Professional" /Compress:max
dism /Export-Image /SourceImageFile:image\ULTIMATE.wim /SourceIndex:1 /DestinationImageFile:image\installfinal.wim /DestinationName:"Windows 7 Ultimate" /Compress:max
dism /Export-Image /SourceImageFile:image\HOMEPREMIUM.wim /SourceIndex:1 /DestinationImageFile:image\installfinal.wim /DestinationName:"Windows 7 Home Premium" /Compress:max
dism /Export-Image /SourceImageFile:image\HOMEBASIC.wim /SourceIndex:1 /DestinationImageFile:image\installfinal.wim /DestinationName:"Windows 7 Home Basic" /Compress:max
dism /Export-Image /SourceImageFile:image\STARTER.wim /SourceIndex:1 /DestinationImageFile:image\installfinal.wim /DestinationName:"Windows 7 Starter" /Compress:max

:: Split WIM into SWM
TITLE Splitting WIM
Dism /Split-Image /ImageFile:image\installfinal.wim /SWMFile:image\install.swm /FileSize:4700
IF EXIST %ORIGINALISO%\sources\install.wim MOVE %ORIGINALISO%\sources\install.wim image\installoriginal.wim
COPY Image\*.swm %ORIGINALISO%\sources

:: Create ISO
::  Remove original WIM
TITLE Creating ISO
IF EXIST Windows_7_SP1_AIO_Final.ISO DEL Windows_7_SP1_AIO_Final.ISO
oscdimg.exe -lWindows_7_SP1_AIO_Final -m -u2 -b%ORIGINALISO%\Boot\etfsboot.com %ORIGINALISO% Windows_7_SP1_AIO_Final.ISO
TITLE DONE