' By http://kevinisms.fason.org
'
'*******************************************************************************
'Get Processor Info from WMI
'*******************************************************************************
' Create global variable
Dim gsProcArch

' Connect to WMI Namespace
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2") 

' Connect to Win32_Processor class
Set colItems = objWMIService.ExecQuery("Select * from Win32_Processor") 

For each objItem in colItems
	gsProcArch = Trim(objItem.DataWidth)
next

' setup link to OSD and create OSD variable
Set oEnv = CreateObject("Microsoft.SMS.TSEnvironment")
oEnv("OSDProcArch") = gsProcArch
