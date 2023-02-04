' From http://kevinisms.fason.org
' Written by Cory Becht
' use "cscript PinnedItems.vbs RECORD" to capture pinnings
' use "cscript PinnedItems.vbs ACTIVESETUP" to restore pinnings 

strScriptVersion = "2"
Set filesys = CreateObject("Scripting.FileSystemObject")
Set wshshell = CreateObject("WScript.Shell")
Set objShell = CreateObject("Shell.Application") 
Set objMSI = CreateObject("WindowsInstaller.Installer")
Set strTemp = filesys.GetSpecialFolder(2)
Set strSYS32 = filesys.GetSpecialFolder(1)
Set strWin = filesys.GetSpecialFolder(0)
strPinnedDataFileName = ScriptPath()&"\PinnedItems.dat"
Const strHKLM = &H80000002
Const strHKCU = &H80000001

strVer = GetWindowsVersion()
strx = CStr(CDbl(1/2))
GetDecimalChar = Mid(strx, 2, 1)
If GetDecimalChar <> "." Then
	strVer = Replace(strVer, ".", GetDecimalChar)	
End If
dblWindowsVer = cdbl(strVer)

If ucase(WScript.Arguments(0)) = "LAUNCH" Then
	wshshell.Run strSYS32&"\cscript.exe"&" "&Chr(34)&ScriptPath()&"\"&WScript.ScriptName&Chr(34)&" RESTORE "&Chr(34)&WScript.Arguments(1)&Chr(34), 0, False
End If

If ucase(WScript.Arguments(0)) = "ACTIVESETUP" Then
	Call CreateActiveSetup()
End If

If ucase(WScript.Arguments(0)) = "RECORD" Then
	If filesys.FileExists(strPinnedDataFileName) Then
	   filesys.DeleteFile strPinnedDataFileName, True 
	End If 
	intCount = 0
	ReDim arrLNKFiles(intCount)
	arrProfileLocations = GetProfiles()
	If IsNull(arrProfileLocations) = True Then
		WScript.Quit(0)
	End If
	For Each strProfilePath In arrProfileLocations
		strStartMenuPath = strProfilePath&"\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\StartMenu"
		Call GetFiles(strProfilePath,strStartMenuPath)
		strTaskBarPath = strProfilePath&"\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar"
		Call GetFiles(strProfilePath,strTaskBarPath)
		If dblWindowsVer > 6.1 Then
			strStartMenuPath = strProfilePath&"\AppData\Roaming\Microsoft\Windows\Start Menu\Programs"
			Call GetFiles(strProfilePath,strStartMenuPath)	
		End If
	Next
	If arrLNKFiles(0) <> "" Then
		Set objPinnedDataFile = filesys.OpenTextFile(strPinnedDataFileName,8,True)
		For Each strLNKFile In arrLNKFiles
			aLNKInfo = Split(strLNKFile, "|")
			Set objlink = wshshell.CreateShortcut(aLNKInfo(2))
			strTargetPath = objlink.TargetPath
			If InStr(1, strTargetPath, "\Windows\Installer\", 1) > 0 Then
				Set objMSITarget = objMSI.ShortcutTarget(aLNKInfo(2))
				strTargetPath = objMSI.ComponentPath(objMSITarget.StringData(1), objMSITarget.StringData(3))
			End If
			objPinnedDataFile.WriteLine aLNKInfo(0)&"|"&aLNKInfo(1)&"|"&aLNKInfo(2)&"|"&strTargetPath
		Next	
		objPinnedDataFile.Close
	End if		
End If

If ucase(WScript.Arguments(0)) = "RESTORE" Then
	WScript.Sleep 15000
	strPinnedDataFileName = WScript.Arguments(1)
	If filesys.FileExists(strPinnedDataFileName) = False Then
		WScript.Quit(0)
	End If
	Set objPinnedDataFile = filesys.OpenTextFile(strPinnedDataFileName,1,True)
	strCurrentProfile = GetCurrentProfile()
	Do Until objPinnedDataFile.AtEndOfLine
		strLine = objPinnedDataFile.ReadLine
		arrLine = Split(strLine, "|")
		If UCase(strCurrentProfile) = UCase(arrLine(0)) Then
			If filesys.FileExists(arrLine(2)) = False then
				If filesys.FileExists(arrLine(3)) = True Then
					'get startmenu or taskbar
					arrShortcutInfo = Split(arrLine(2), "\")
					strPinLocation = ucase(arrShortcutInfo(UBound(arrShortcutInfo)-1))
					'Create shortcut in profile
					Set objlink = wshshell.CreateShortcut(arrLine(0)&"\"&arrLine(1)&".lnk")
					objLink.TargetPath = arrLine(3)
					objLink.Save
					Call CreatePin(arrLine(0), arrLine(1)&".lnk", strPinLocation)
				Else
					strNewpath = GetNewPath(arrLine(3))
					arrShortcutInfo = Split(arrLine(2), "\")
					strPinLocation = ucase(arrShortcutInfo(UBound(arrShortcutInfo)-1))
					Set link = wshshell.CreateShortcut(arrLine(0)&"\"&arrLine(1)&".lnk")
					link.TargetPath = strNewpath
					link.Save
					Call CreatePin(arrLine(0), arrLine(1)&".lnk", strPinLocation)
				End If
			End If
		End If
	Loop
End If

WScript.Quit(0)
'--------------------------------------------------------------------------------------------------
Sub CreateActiveSetup()
	Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
	sKeyPath = "SOFTWARE\Microsoft\Active Setup\Installed Components\O365PinScript"
	oReg.CreateKey strHKLM, sKeyPath
	sValueName1 = "Version"
	sValueName2 = "StubPath"
	oReg.SetStringValue strHKLM,sKeyPath,sValueName1,strScriptVersion
	oReg.SetStringValue strHKLM,sKeyPath,sValueName2,strSYS32&"\cscript.exe "&Chr(34)&ScriptPath()&"\"&WScript.ScriptName&Chr(34)&" LAUNCH "&Chr(34)&strPinnedDataFileName&Chr(34)
	oReg.SetStringValue strHKLM,sKeyPath,"", "Office 365 Items"
End Sub

Sub CreatePin(sFolder, sLNK, sPinLocation)
	Set objFolder = objShell.Namespace(sFolder)
	Set objFolderItem = objFolder.ParseName(sLNK) 
	Set colVerbs = objFolderItem.Verbs 
	For Each objVerb in colVerbs 
		If sPinLocation = "STARTMENU" Then
  			If ucase(Replace(objVerb.name, "&", "")) = "PIN TO START MENU" or ucase(Replace(objVerb.name, "&", "")) = "PIN TO START" Then 
  				objVerb.DoIt
  				Exit For 
  			End If
 		End If
 		If strPinLocation = "TASKBAR" Then
  			If ucase(Replace(objVerb.name, "&", "")) = "PIN TO TASKBAR" Then 
  				objVerb.DoIt
  				Exit For 
  			End If
 		End If
	Next
End Sub

Function GetCurrentProfile()
	'GetCurrentProfile = wshShell.ExpandEnvironmentStrings( "%USERPROFILE%" )
	Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
	sKeyPath4 = "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
	sValueName4 = "Desktop"
	oReg.GetStringValue strHKCU,sKeyPath4,sValueName4,sValue4
	iLoc = InStrRev(sValue4, "\Desktop")
	sProf = Left(sValue4, iLoc-1) 
	GetCurrentProfile = sProf
End Function

Function GetNewPath(sFile)
	aFile = Split(sFile, "\")
	sFilePath = ""
	For i = 0 To UBound(aFile)
		If UCase(aFile(i)) = "OFFICE15" Then
			iStart = i + 1
			Exit For
		End If
	Next
	For i = iStart To UBound(aFile)
		sFilePath = sFilePath &"\"&aFile(i)
	Next
	Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
	sKeyPath = "SOFTWARE\Microsoft\Office\15.0\ClickToRun\Configuration"
	sValueName = "InstallationPath"
	oReg.GetStringValue strHKLM,sKeyPath,sValueName,sValue
	GetNewPath = sValue & "\root\office15"&sFilePath
End Function

Sub GetFiles(sProfile,sPath)
		If filesys.FolderExists(sPath) = True Then
			For Each oFile In filesys.GetFolder(sPath).Files
				If UCase(filesys.GetExtensionName(oFile.Name)) = "LNK" Then
					ReDim Preserve arrLNKFiles(intCount)
					arrLNKFiles(intCount) = sProfile&"|"&Left(oFile.Name,Len(oFile.Name)-4)&"|"&oFile.Path
					intCount = intCount + 1
				End If
			Next
		End If
End Sub

Function GetProfiles()
	iCount = 0
	bln1 = False
	Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
 	sKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
	oReg.EnumKey strHKLM, sKeyPath, aSubKeys
 	For Each sSubkey In aSubKeys
    	 If Len(sSubkey) > 8 Then
    	 	sKey1 = sKeyPath & "\" & sSubkey
    	 	sValueName = "ProfileImagePath"
    	 	oReg.GetExpandedStringValue strHKLM, sKey1, sValueName, sValue1
    	 	ReDim Preserve aProfs(iCount)
    	 	aProfs(iCount) = sValue1    	 
    	 	iCount = iCount + 1
    	 	bln1 = True
    	 End If
	Next
	If bln1 = True Then
		GetProfiles = aProfs
	Else
		GetProfiles = vbNull
	End If
End Function

Function GetWindowsVersion()
	strVer = wshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\CurrentVersion")
	GetWindowsVersion = strVer
End Function

Function ScriptPath()
	ScriptPath = Left(WScript.ScriptFullName, Len(WScript.ScriptFullName) - Len(WScript.ScriptName)-1)
End Function
