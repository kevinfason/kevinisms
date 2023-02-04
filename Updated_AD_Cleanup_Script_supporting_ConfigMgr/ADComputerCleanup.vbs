' by http://kevinisms.fason.org
'  written by Cory Becht
'Max days of computer not logging into AD.  if a computer is older than this number it will be moved into quarantine
Const intMaxDaysOldQurantine = 90
'Max days to stay in quarantine before disabling account
Const intMaxDaysInQuarantine = 30
'Max days to stay in quarantine before purging from AD
Const intMaxDaysOldPurge = 30
'Exclude Servers
Const blnExcludeServers = True
'SMTP server to send email to
Const strSMTP = "smtpsmarthost.mydomain.local"
'SMTP return email address
Const STRFROM = "ADCleanupScript@mydomain.local"
'SMTP email address to send notifications
Const StrTo = "recipients@mydomain.local,recipients2@mydomain.local"
'Whether or not to just run a test.  This will not move accounts or delete anything
Const blnTestOnly = False

'Workstation OS names to move
arrWorkstations = Array("XP","Vista","Windows 7","Windows 8", "Windows 8.1", "Windows 10")

'Specific Computer Names to exclude
GblstrExcludedComputersFile = ScriptPath()&"\ExcludedComputers.txt"
'Specific OU names to exclude.
GblstrExcludedOUsFile = ScriptPath()&"\ExcludedOUs.txt"
'Location to place moved computer accounts
GblstrComputerQuarantineStart = "ou=ADCleanup"
GblstrComputerQuarantineActive = "ou=Active,ou=ADCleanup"
GblstrComputerQuarantineDisabled = "ou=Disabled,ou=ADCleanup"

'ConfigMgr info
Const strSCCMSQL = "ConfigMgr.mydomain.local"
Const strSCCMDB = "CM_SITECODE"

Set filesys = CreateObject("Scripting.FileSystemObject")
Set rsData = CreateObject("ADOR.Recordset")
Set rsMoveList = CreateObject("ADOR.Recordset")
Set wshshell = CreateObject("WScript.Shell")
Const adVarChar = 200
Const adVarWChar = 202
Const adDBTimeStamp = 135
Const adDouble = 5
Const adBoolean = 11
Const MaxCharacters = 255
GblstrWorkstationEmailMessage = "Starting Old Computer Cleanup Script"&vbCrLf&vbCrLf
GblstrServerEmailMessage = "Starting Old Computer Cleanup Script"&vbCrLf&vbCrLf
If blnTestOnly = True Then
	GblstrWorkstationEmailMessage = GblstrWorkstationEmailMessage & "******************* This is only a test run.  No changes have been made. ***********************"&vbCrLf
	GblstrWorkstationEmailMessage = GblstrWorkstationEmailMessage & vbCrLf
	GblstrServerEmailMessage = GblstrServerEmailMessage & "******************* This is only a test run.  No changes have been made. ***********************"&vbCrLf
	GblstrServerEmailMessage = GblstrServerEmailMessage & vbCrLf
End if
lngTimeBias = GetTimeZoneBias()
Set objRootDSE = GetObject("LDAP://RootDSE")
strDefaultNamingContext = objRootDSE.Get("defaultNamingContext")

'Get Excluded Computers List
arrExcludedComputers = ConvertTextToArray(GblstrExcludedComputersFile)
'Get Excluded OUs List
arrExcludedOUs = ConvertTextToArray(GblstrExcludedOUsFile)
'Purge any old accounts from the quarantine area
Call PurgeOld()
If GblstrWorkstationEmailMessage <> "" Then 
	GblstrWorkstationEmailMessage = GblstrWorkstationEmailMessage & vbCrLf
End If
If GblstrServerEmailMessage <> "" Then 	
	GblstrServerEmailMessage = GblstrServerEmailMessage & vbCrLf
End If
'Get info about computer accounts into recordset
gblMoveandDisable = False
Call GetComputerObjectInfo()
If GblstrWorkstationEmailMessage <> "" Then 
	GblstrWorkstationEmailMessage = GblstrWorkstationEmailMessage & vbCrLf
End If
If GblstrServerEmailMessage <> "" Then 	
	GblstrServerEmailMessage = GblstrServerEmailMessage & vbCrLf
End If
'disable the computer accounts that have not updated the last logon timestamp within intMaxDaysInQuarantine days and are in the quarantine ou
Call Disable()
If GblstrWorkstationEmailMessage <> "" Then 
	GblstrWorkstationEmailMessage = GblstrWorkstationEmailMessage & vbCrLf
End If
If GblstrServerEmailMessage <> "" Then 	
	GblstrServerEmailMessage = GblstrServerEmailMessage & vbCrLf
End If
If gblMoveandDisable = True Then
	'Move the computer accounts that have not logged in with value specified within intMaxDaysOldQuarantine variable
	Call Move()
	If GblstrWorkstationEmailMessage <> "" Then 
		GblstrWorkstationEmailMessage = GblstrWorkstationEmailMessage & vbCrLf
	End If
	If GblstrServerEmailMessage <> "" Then 	
		GblstrServerEmailMessage = GblstrServerEmailMessage & vbCrLf
	End If
End If
'Move any enabled accounts back to their original location.
Call MoveBack(GblstrComputerQuarantineActive,True)
Call MoveBack(GblstrComputerQuarantineDisabled,False)

GblstrWorkstationEmailMessage = GblstrWorkstationEmailMessage & vbCrLf& "Ending Old Computer Cleanup Script"
GblstrServerEmailMessage = GblstrServerEmailMessage & vbCrLf & "Ending Old Computer Cleanup Script"

If GblstrWorkstationEmailMessage <> "" Then
	Call Email("Workstation Accounts Cleanup Script Ran", GblstrWorkstationEmailMessage)
End If
If GblstrServerEmailMessage <> "" Then
	Call Email("Server Accounts Cleanup Script Ran", GblstrServerEmailMessage)
End If

WScript.Quit(0)
'--------------------------------------------
Function RemoveUnicode(strString)
	strT = ""
	intLen = Len(strString)
	For i = 1 To intLen
		If Asc(Mid(strString, i, 1)) >= 0 And Asc(Mid(strString, i, 1)) <= 239 And Asc(Mid(strString, i, 1)) <> 63 Then
			strT = strT & Mid(strString, i, 1)
		End If	
	Next
	RemoveUnicode = strT	
End Function

Sub PurgeOld()
	Const ADS_UF_ACCOUNTDISABLE=2
	Set objConn = CreateObject("ADODB.Connection")
	objConn.Open "Provider=ADsDSOObject"
	Set objCommand =CreateObject("ADODB.Command")
	objCommand.ActiveConnection = objConn
	objCommand.Properties("Page Size") = 99999
	objCommand.Properties("Chase referrals") = &H60
	objCommand.CommandText = "<LDAP://"&GblstrComputerQuarantineDisabled&","&strDefaultNamingContext&">;(objectCategory=computer);description,userAccountControl,distinguishedName,cn,operatingSystem;subtree"
	Set objRecordSet = objCommand.Execute
	Do Until objRecordSet.EOF
		intUAC = objRecordSet.Fields("userAccountControl").value 
		If intUAC And ADS_UF_ACCOUNTDISABLE Then
			arrDesc = objRecordSet.Fields("description").value
			If VarType(arrDesc) <> 8204 Then
				strDescription = ""
			Else
				strDescription = arrDesc(0)
			End If
			If strDescription <> "" Then 
				If InStr(1, strDescription, "::", 1) > 0 Then
					If InStr(1, strDescription, "[", 1) > 0 Then
						If InStr(1, strDescription, "]", 1) > 0 Then
							dtDisabledDate = GetTime(strDescription,objRecordSet.Fields("distinguishedName").value)
							If DateDiff("d", dtDisabledDate, Now ) > intMaxDaysOldPurge Then
								'Delete Account
								If blnTestOnly = False Then
									Set objComputer = GetObject("LDAP://"&objRecordSet.Fields("distinguishedName").value)
									objComputer.deleteobject (0)
								End If
								If InStr(1, objRecordSet.Fields("operatingSystem").value, "Server", 1) > 0 Then
									GblstrServerEmailMessage = GblstrServerEmailMessage & "Deleted computer account "&objRecordset.Fields("cn").value& ". Account was disabled on "& dtDisabledDate &"."&vbCrLf
								Else
									GblstrWorkstationEmailMessage = GblstrWorkstationEmailMessage & "Deleted computer account "&objRecordset.Fields("cn").value& ".  Account was disabled on "& dtDisabledDate &"."&vbCrLf
								End If
							End If							
						Else
							If InStr(1, objRecordSet.Fields("operatingSystem").value, "Server", 1) > 0 Then
								GblstrServerEmailMessage = GblstrServerEmailMessage & "Disabled computer account "&objRecordset.Fields("cn").value& " does not have correct disabled time stamp."& vbCrLf
							Else
								GblstrWorkstationEmailMessage = GblstrWorkstationEmailMessage & "Disabled computer account "&objRecordset.Fields("cn").value& " does not have correct disabled time stamp."& vbCrLf
							End If
						End If
					Else
						If InStr(1, objRecordSet.Fields("operatingSystem").value, "Server", 1) > 0 Then
							GblstrServerEmailMessage = GblstrServerEmailMessage & "Disabled computer account "&objRecordset.Fields("cn").value& " does not have correct disabled time stamp."& vbCrLf
						Else
							GblstrWorkstationEmailMessage = GblstrWorkstationEmailMessage & "Disabled computer account "&objRecordset.Fields("cn").value& " does not have correct disabled time stamp."& vbCrLf
						End If
					End If
				Else
					If InStr(1, objRecordSet.Fields("operatingSystem").value, "Server", 1) > 0 Then
						GblstrServerEmailMessage = GblstrServerEmailMessage & "Disabled computer account "&objRecordset.Fields("cn").value& " does not have correct disabled time stamp."& vbCrLf
					Else
						GblstrWorkstationEmailMessage = GblstrWorkstationEmailMessage & "Disabled computer account "&objRecordset.Fields("cn").value& " does not have correct disabled time stamp."& vbCrLf
					End If
				End If				
			Else
				If InStr(1, objRecordSet.Fields("operatingSystem").value, "Server", 1) > 0 Then
					GblstrServerEmailMessage = GblstrServerEmailMessage & "Disabled computer account "&objRecordset.Fields("cn").value& " does not have a description set.  Cannot read the disabled timestamp."& vbCrLf
				Else
					GblstrWorkstationEmailMessage = GblstrWorkstationEmailMessage & "Disabled computer account "&objRecordset.Fields("cn").value& " does not have a description set.  Cannot read the disabled timestamp."& vbCrLf
				End If
			End If
		End If	
		objRecordSet.Movenext
	Loop
End Sub

Function GetTime(strDesc, sObj)
	'WScript.Echo sObj&"|"&strDesc
	intLoc1 = InStr(1, strDesc, "[", 1)
	intLoc2 = InStr(1, strDesc, "]", 1)
	If intLoc1 <> 0 And intLoc2 <> 0 Then
		GetTime = cdate(Trim(Mid(strDesc, intLoc1 + 1, (intLoc2-intLoc1)-1)))
	Else
		GetTime = Now()
		Call SetDescription(sObj, 1)
	End if
End Function

Sub MoveBack(strQOU, blnCheckLastLogonTimeStamp)
	ADS_UF_ACCOUNTDISABLE=2
	Set objConn = CreateObject("ADODB.Connection")
	objConn.Open "Provider=ADsDSOObject"
	Set objCommand =CreateObject("ADODB.Command")
	objCommand.ActiveConnection = objConn
	objCommand.Properties("Page Size") = 99999
	objCommand.Properties("Chase referrals") = &H60
	objCommand.CommandText = "<LDAP://"&strQOU&","&strDefaultNamingContext&">;(objectCategory=computer);userAccountControl,distinguishedName,cn,desktopProfile,lastLogonTimeStamp,description,operatingSystem;subtree"
	Set objRecordSet = objCommand.Execute
	Do Until objRecordSet.EOF
		intUAC = objRecordSet.Fields("userAccountControl").value 
		If Not intUAC And ADS_UF_ACCOUNTDISABLE Then
			blnCont = True
			If blnCheckLastLogonTimeStamp = True Then
				On Error Resume Next
				Set objLastLogon = objRecordSet.Fields("lastLogonTimeStamp").value
				If Err.Number = 0 Then
					dtLastLogon = ConvertTime(objLastLogon)
				Else
					dtLastLogon = #1/1/1601#
				End If
				On Error GoTo 0
				If IsNull(objRecordSet.Fields("description").value) Then
					strDescription = ""
				Else
					arrDesc = objRecordSet.Fields("description").value
					strDescription = arrDesc(0)
				End if
				'arrDesc = objRecordSet.Fields("description").value
  				'strDescription = arrDesc(0)
				If dtLastLogon < GetTime(strDescription, objRecordSet.Fields("distinguishedName").value) Then
					blnCont = False
				End If
			End If
			If blnCont = True Then
				If objRecordset.Fields("desktopProfile").value <> "" And objRecordset.Fields("desktopProfile").value <> vbNull Then
					intReturnMove = 0
					If blnTestOnly = False Then
						intReturnMove = MoveComputer(objRecordset.Fields("cn").value,objRecordset.Fields("distinguishedName").value,objRecordset.Fields("desktopProfile").value,2)
					End If
					If intReturnMove = 0 Then
						intReturnDesc = 0
						If blnTestOnly = False Then
							intReturnDesc = SetDescription("CN="&objRecordset.Fields("cn").value&","&objRecordset.Fields("desktopProfile").value,2)
						End If
						If intReturnDesc = 0 Then
							If Instr(1,objRecordSet.Fields("operatingSystem").value,"Server",1) > 0 Then
								GblstrServerEmailMessage = GblstrServerEmailMessage & "Enabled computer account "&objRecordset.Fields("cn").value& " was moved back to OU location of "&objRecordset.Fields("desktopProfile").value&"." & vbCrLf
							Else
								GblstrWorkstationEmailMessage = GblstrWorkstationEmailMessage & "Enabled computer account "&objRecordset.Fields("cn").value& " was moved back to OU location of "&objRecordset.Fields("desktopProfile").value&"." & vbCrLf
							End If
						Else
						
						End if					
					Else
						If intReturnMove = 424 Then
							If Instr(1,objRecordSet.Fields("operatingSystem").value,"Server",1) > 0 Then
								GblstrServerEmailMessage = GblstrServerEmailMessage & "Enabled computer account "&objRecordset.Fields("cn").value& " has previous OU location of "&objRecordset.Fields("desktopProfile").value&" but it does not exist and must be moved manually to the correct location." & vbCrLf
							Else
								GblstrWorkstationEmailMessage = GblstrWorkstationEmailMessage & "Enabled computer account "&objRecordset.Fields("cn").value& " has previous OU location of "&objRecordset.Fields("desktopProfile").value&" but it does not exist and must be moved manually to the correct location." & vbCrLf
							End If
						Else 
							If Instr(1,objRecordSet.Fields("operatingSystem").value,"Server",1) > 0 Then
								GblstrServerEmailMessage = GblstrServerEmailMessage & "Enabled computer account "&objRecordset.Fields("cn").value& " has previous OU location of "&objRecordset.Fields("desktopProfile").value&" but it could not be moved and must be moved manually to the correct location.  Error:" &intReturn& vbCrLf
							Else
								GblstrWorkstationEmailMessage = GblstrWorkstationEmailMessage & "Enabled computer account "&objRecordset.Fields("cn").value& " has previous OU location of "&objRecordset.Fields("desktopProfile").value&" but it could not be moved and must be moved manually to the correct location.  Error:" &intReturn& vbCrLf
							End If
						End If
					End if				
				Else
					If Instr(1,objRecordSet.Fields("operatingSystem").value,"Server",1) > 0 Then
						GblstrServerEmailMessage = GblstrServerEmailMessage & "Enabled computer account "&objRecordset.Fields("cn").value& " does not have previous OU location set and must be moved manually to the correct location." & vbCrLf
					Else
						GblstrWorkstationEmailMessage = GblstrWorkstationEmailMessage & "Enabled computer account "&objRecordset.Fields("cn").value& " does not have previous OU location set and must be moved manually to the correct location." & vbCrLf
					End If
				End If
			End If
		End If	
		objRecordSet.Movenext
	Loop
End Sub

Function GetSCCMUDAInfo(sComputer)
	sHW = "Not Found"
	sUDA = "Not Found"
	Set OConn = CreateObject("ADODB.Connection")
	oConn.Open "Driver={SQL Server};Server="&strSCCMSQL&";Database="&strSCCMDB&";Trusted_Connection=yes"
	sQuery = "select distinct WS.LastHWScan, RUSER.Unique_User_Name0 " _
			& "from v_R_System SYS " _
			& "outer apply (select top 1 * from v_UserMachineRelation UMR1 " _
			& "where UMR1.MachineResourceID = sys.ResourceID and UMR1.RelationActive = 1 " _
			& "order by rowversion DESC) as UMR " _
			& "LEFT JOIN v_R_User RUSER on UMR.UniqueUserName=RUSER.Unique_User_Name0 " _
			& "LEFT JOIN v_GS_Workstation_Status WS on SYS.ResourceID=WS.ResourceID " _
			& "WHERE SYS.Name0 = '"& sComputer&"'"
	Set oRst = oConn.Execute(sQuery,,adCmdTxt)
	While Not oRst.eof
		sUDA = ORst.Fields("Unique_User_Name0").value
		If IsNull(sUDA) = True Then
			sUDA = "No Info Available"
		End If
		sHW = ORst.Fields("LastHWScan").value
		If IsNull(sHW) = True Then
			sHW = "No Info Available"
		End If
		oRst.MoveNext
	Wend
	OConn.Close
	GetSCCMUDAInfo = sUDA&"|"&sHW
End Function

Sub Disable()
	ADS_UF_ACCOUNTDISABLE=2
	Set objConn = CreateObject("ADODB.Connection")
	objConn.Open "Provider=ADsDSOObject"
	Set objCommand =CreateObject("ADODB.Command")
	objCommand.ActiveConnection = objConn
	objCommand.Properties("Page Size") = 99999
	objCommand.Properties("Chase referrals") = &H60
	objCommand.CommandText = "<LDAP://"&GblstrComputerQuarantineActive&","&strDefaultNamingContext&">;(objectCategory=computer);distinguishedName,cn,description,operatingSystem;subtree"
	Set objRecordSet = objCommand.Execute
	Do Until objRecordSet.EOF
		'WScript.Echo objRecordSet.Fields("distinguishedName").value
		'WScript.Echo objRecordSet.Fields("description").value
		If IsNull(objRecordSet.Fields("description").value) Then
			strDescription = ""
		Else
			arrDesc = objRecordSet.Fields("description").value
			strDescription = arrDesc(0)
		End if
		If strDescription <> "" Then 
			If InStr(1, strDescription, "::", 1) > 0 Then
				If InStr(1, strDescription, "[", 1) > 0 Then
					If InStr(1, strDescription, "]", 1) > 0 Then
						dtMovedDate = GetTime(strDescription, objRecordSet.Fields("distinguishedName").value)
						If DateDiff("d", dtMovedDate, Now ) > intMaxDaysInQuarantine Then
							'disable account
							'get UDA info from SCCM
							sSCCMInfo = GetSCCMUDAInfo(objRecordset.Fields("cn").value)
							aSCCMInfo = Split(sSCCMInfo,"|")							
							intReturnDisable = 0
							If blnTestOnly = False Then
								Set objComputer = GetObject("LDAP://"&objRecordSet.Fields("distinguishedName").value)
								intUAC = objComputer.Get("userAccountControl")
								objComputer.Put "userAccountControl", intUAC Or ADS_UF_ACCOUNTDISABLE
								intReturnDisable = objComputer.setinfo
							End If
							If intReturnDisable = 0 Then
								If Instr(1,objRecordSet.Fields("operatingSystem").value,"Server",1) > 0 Then
									GblstrServerEmailMessage = GblstrServerEmailMessage & "Disabled computer account "&objRecordset.Fields("cn").value& ". Last SCCM Inventory:"&aSCCMInfo(1)&". Account was moved on "& dtMovedDate &"."&vbCrLf
								Else
									GblstrWorkstationEmailMessage = GblstrWorkstationEmailMessage & "Disabled computer account "&objRecordset.Fields("cn").value& ".  Last SCCM Inventory:"&aSCCMInfo(1)&". Primary User:"&aSCCMInfo(0)&". Account was moved on "& dtMovedDate &". Description: "&strDescription&vbCrLf
								End If
								'move account
								intReturnMove = 0
								strNewDN = "CN="&objRecordSet.Fields("cn").value&","&GblstrComputerQuarantineDisabled&","&strDefaultNamingContext
								If blnTestOnly = False Then	
									Set objComputer = GetObject("LDAP://"&objRecordSet.Fields("distinguishedName").value)
									intReturnMove = MoveComputer(objRecordSet.Fields("cn").value,objRecordSet.Fields("distinguishedName").value,GblstrComputerQuarantineDisabled&","&strDefaultNamingContext,0)
								End If
								If intReturnMove = 0 Then
									If Instr(1,objRecordSet.Fields("operatingSystem").value,"Server",1) > 0 Then
										GblstrServerEmailMessage = GblstrServerEmailMessage &vbTab&objRecordSet.Fields("cn").value&" was moved to "&GblstrComputerQuarantineDisabled&","&strDefaultNamingContext&"."&vbCrLf
									Else 
										GblstrWorkstationEmailMessage = GblstrWorkstationEmailMessage &vbTab&objRecordSet.Fields("cn").value&" was moved to "&GblstrComputerQuarantineDisabled&","&strDefaultNamingContext&"."&vbCrLf
									End If
									If blnTestOnly = False Then	
										intReturnDesc = SetDescription(strNewDN,3)
									End If
									If intReturnDesc = 0 Then
										If Instr(1,objRecordSet.Fields("operatingSystem").value,"Server",1) > 0 Then
											GblstrServerEmailMessage = GblstrServerEmailMessage &vbTab&"Updated the description for account "&objRecordSet.Fields("cn").value&"."&vbCrLf
										Else
											GblstrWorkstationEmailMessage = GblstrWorkstationEmailMessage &vbTab&"Updated the description for account "&objRecordSet.Fields("cn").value& "."&vbCrLf
										End If
									Else
										If Instr(1,objRecordSet.Fields("operatingSystem").value,"Server",1) > 0 Then
											GblstrServerEmailMessage = GblstrServerEmailMessage &vbTab& "Unable to update the description for account "&objRecordSet.Fields("cn").value& " Error:"&intReturnDisable&".  Account will remain disabled."&vbCrLf
										Else
											GblstrWorkstationEmailMessage = GblstrWorkstationEmailMessage &vbTab& "Unable to update the description for account "&objRecordSet.Fields("cn").value& " Error:"&intReturnDisable&".  Account will remain disabled."&vbCrLf
										End If
									End If
								Else
									If Instr(1,objRecordSet.Fields("operatingSystem").value,"Server",1) > 0 Then
										GblstrServerEmailMessage = GblstrServerEmailMessage & "Unable to move computer account "&objRecordset.Fields("cn").value& "  to "&GblstrComputerQuarantineDisabled&" Error:"& intReturnMove &"."&vbCrLf
									Else
										GblstrWorkstationEmailMessage = GblstrWorkstationEmailMessage & "Unable to move computer account "&objRecordset.Fields("cn").value& "  to "&GblstrComputerQuarantineDisabled&" Error:"& intReturnMove &"."&vbCrLf
									End If
								End If	
							Else
								If Instr(1,objRecordSet.Fields("operatingSystem").value,"Server",1) > 0 Then
									GblstrServerEmailMessage = GblstrServerEmailMessage & "Unable to disable account "&objRecordset.Fields("cn").value& " Error:"& intReturnDisable &"."&vbCrLf
								Else 
									GblstrWorkstationEmailMessage = GblstrWorkstationEmailMessage & "Unable to disable account "&objRecordset.Fields("cn").value& " Error:"& intReturnDisable &"."&vbCrLf
								End If
							End If
						End If							
					Else
						If Instr(1,objRecordSet.Fields("operatingSystem").value,"Server",1) > 0 Then
							GblstrServerEmailMessage = GblstrServerEmailMessage & "Computer account "&objRecordset.Fields("cn").value& " does not have correct moved time stamp."& vbCrLf
						Else
							GblstrWorkstationEmailMessage = GblstrWorkstationEmailMessage & "Computer account "&objRecordset.Fields("cn").value& " does not have correct moved time stamp."&vbCrLf
						End If
					End If
				Else
					If Instr(1,objRecordSet.Fields("operatingSystem").value,"Server",1) > 0 Then
						GblstrServerEmailMessage = GblstrServerEmailMessage & "Computer account "&objRecordset.Fields("cn").value& " does not have correct moved time stamp."&vbCrLf
					Else
						GblstrWorkstationEmailMessage = GblstrWorkstationEmailMessage & "Computer account "&objRecordset.Fields("cn").value& " does not have correct moved time stamp."&vbCrLf
					End If
				End If
			Else
				If Instr(1,objRecordSet.Fields("operatingSystem").value,"Server",1) > 0 Then
					GblstrServerEmailMessage = GblstrServerEmailMessage & "Computer account "&objRecordset.Fields("cn").value& " does not have correct moved time stamp."&vbCrLf
				Else
					GblstrWorkstationEmailMessage = GblstrWorkstationEmailMessage & "Computer account "&objRecordset.Fields("cn").value& " does not have correct moved time stamp."&vbCrLf
				End If
			End If				
		Else
			If Instr(1,objRecordSet.Fields("operatingSystem").value,"Server",1) > 0 Then
				GblstrServerEmailMessage = GblstrServerEmailMessage & "Computer account "&objRecordset.Fields("cn").value& " does not have a description set.  Cannot read the moved timestamp."&vbCrLf
			Else
				GblstrWorkstationEmailMessage = GblstrWorkstationEmailMessage & "Computer account "&objRecordset.Fields("cn").value& " does not have a description set.  Cannot read the moved timestamp."&vbCrLf
			End If
		End If
		objRecordSet.Movenext
	Loop
End Sub

Sub Move()
	rsData.Movefirst
	rsData.Sort = "DaysOld DESC"
	rsMoveList.Fields.Append "Computer",adVarChar,MaxCharacters
	rsMoveList.Fields.Append "DaysOld",adDouble,MaxCharacters
	rsMoveList.Fields.Append "OldDN",adVarChar,MaxCharacters
	rsMoveList.Fields.Append "WhenCreated",adDBTimeStamp,MaxCharacters
	rsMoveList.Fields.Append "Success",adBoolean,MaxCharacters
	rsMoveList.Fields.Append "OS",adVarWChar,MaxCharacters
	rsMoveList.Fields.Append "Description",adVarWChar,MaxCharacters
	rsMoveList.Open
	Do While Not rsData.EOF
		If rsData.Fields.Item("DaysOld").value > intMaxDaysOldQurantine Then
			If InStr(1, rsData.Fields.Item("DN").value, GblstrComputerQuarantineActive&","&strDefaultNamingContext, 1) = 0 Then
				intReturnMove = 0 
				'get UDA info from SCCM
				sSCCMInfo = GetSCCMUDAInfo(rsData.Fields.Item("Computer").value)
				aSCCMInfo = Split(sSCCMInfo,"|")	
				If blnTestOnly = False Then
					intReturnMove = MoveComputer(rsData.Fields.Item("Computer").value,rsData.Fields.Item("DN").value,GblstrComputerQuarantineActive&","&strDefaultNamingContext,1)
				End If
				If intReturnMove = 0 Then
					If blnTestOnly = False Then
						rsMoveList.AddNew
						rsMoveList("OldDN") = rsData.Fields.Item("DN").value
						rsMoveList.Update
					End If
					strNewDN = "CN="&rsData.Fields.Item("Computer").value&","&GblstrComputerQuarantineActive&","&strDefaultNamingContext
					If Instr(1,rsData.Fields.Item("OS").value,"Server",1) > 0 Then
						GblstrServerEmailMessage = GblstrServerEmailMessage & "Moved "&rsData.Fields.Item("Computer").value&" with last logon of "&rsData.Fields.Item("LastLogon").value&" from "&rsData.Fields.Item("DN").value&" to "&GblstrComputerQuarantineActive&","&strDefaultNamingContext&"."& vbCrLf
					Else
						GblstrWorkstationEmailMessage = GblstrWorkstationEmailMessage & "Moved "&rsData.Fields.Item("Computer").value&" with last logon of "&rsData.Fields.Item("LastLogon").value&" from "&rsData.Fields.Item("DN").value&" to "&GblstrComputerQuarantineActive&","&strDefaultNamingContext&". Last SCCM Inventory:"&aSCCMInfo(1)&". Primary User:"&aSCCMInfo(0)&". Description: "&rsData.Fields.Item("Description").value&vbCrLf
					End If
					intReturnDesc = 0
					If blnTestOnly = False Then
						intReturnDesc = SetDescription(strNewDN,1)
					End If
					If intReturnDesc = 0 Then
						If Instr(1,rsData.Fields.Item("OS").value,"Server",1) > 0 Then
							GblstrServerEmailMessage = GblstrServerEmailMessage &vbTab&"Updated the description for account "&rsData.Fields.Item("Computer").value&"."&vbCrLf
						Else
							GblstrWorkstationEmailMessage = GblstrWorkstationEmailMessage &vbTab&"Updated the description for account "&rsData.Fields.Item("Computer").value&"."&vbCrLf
						End If
					Else
						If Instr(1,rsData.Fields.Item("OS").value,"Server",1) > 0 Then
							GblstrServerEmailMessage = GblstrServerEmailMessage &vbTab&"Unable to update the description for account "&rsData.Fields.Item("Computer").value& " Error:"&intReturnDisable&".  Account will remain moved." & vbCrLf
						Else
							GblstrWorkstationEmailMessage = GblstrWorkstationEmailMessage &vbTab&"Unable to update the description for account "&rsData.Fields.Item("Computer").value& " Error:"&intReturnDisable&".  Account will remain moved." & vbCrLf
						End If
					End If							
				Else
					If Instr(1,rsData.Fields.Item("OS").value,"Server",1) > 0 Then
						GblstrServerEmailMessage = GblstrServerEmailMessage & "Unable to move computer account "&rsData.Fields.Item("Computer").value& " from "&rsData.Fields.Item("DN").value&" to "&GblstrComputerQuarantineActive&" Error:"&intReturnMove&"." & vbCrLf
					Else
						GblstrWorkstationEmailMessage = GblstrWorkstationEmailMessage & "Unable to move computer account "&rsData.Fields.Item("Computer").value& " from "&rsData.Fields.Item("DN").value&" to "&GblstrComputerQuarantineActive&" Error:"&intReturnMove&"." & vbCrLf
					End If
				End If
			End If
		End If
	rsData.MoveNext
	Loop
End Sub

'Sets the description for the specified object
Function SetDescription(strObject,intType)
	On Error Resume Next
	Set objComputer = GetObject("LDAP://"&strObject)
	strCurrentDescription = objComputer.description
	If intType = 1 Then
		If InStr(1, strCurrentDescription, "::", 1) = 0 Then
			objComputer.description = objComputer.description&"::Account Automatically Moved - ["&Now&"]"
		Else
			intLoc = InStr(1, strCurrentDescription, "::", 1)
			strDesc = Left(objComputer.description, intLoc - 1)
			objComputer.description = strDesc &"::Account Automatically Disabled - ["&Now&"]"
		End If
	End If
	If intType = 2 Then
		If InStr(1, strCurrentDescription, "::", 1) <> 0 Then
			intLoc = InStr(1, strCurrentDescription, "::", 1)
			strDesc = Left(objComputer.description, intLoc - 1)
			objComputer.description = strDesc
		End If
	End If	
	If intType = 3 Then
		If InStr(1, strCurrentDescription, "::", 1) = 0 Then
			objComputer.description = objComputer.description&"::Account Automatically Disabled - ["&Now&"]"
		Else
			intLoc = InStr(1, strCurrentDescription, "::", 1)
			strDesc = Left(objComputer.description, intLoc - 1)
			objComputer.description = strDesc &"::Account Automatically Disabled - ["&Now&"]"
		End If
	End If
	If intType = 4 Then
		sNewDesc = ""
		iStart = InStr(1,strCurrentDescription, "::Account Automatically ", 1) 
		iEnd = InStr(iStart, strCurrentDescription, "]", 1)
		sNewDesc = Left(strCurrentDescription, iStart - 1)
		sNewDesc = sNewDesc + Mid(strCurrentDescription, iEnd + 1)
		If sNewDesc <> "" Then
			objComputer.description = sNewDesc
		Else
			'set attribute to null
			objComputer.PutEx 1, "description", 0
		End if
	End If
	If blnTestOnly = False Then
		SetDescription = objComputer.setinfo
	End If
	On Error GoTo 0
End Function

'disables the specified account
Function DisableComputer(strObject)
	Set objComputer = GetObject("LDAP://"&strObject)
	objComputer.AccountDisabled = True
	On Error Resume Next
	DisableComputer = objComputer.SetInfo
	On Error GoTo 0
End Function

'Moves the specified accounts from one location to another
Function MoveComputer(strComputer,strObject, strMoveTo, intType)
	On Error Resume Next
	Set objOU = GetObject("LDAP://"&strMoveTo)
	set objReturn = objOU.MoveHere("LDAP://"&strObject, vbNullString)
	If Err.Number <> 0 Then
		MoveComputer = Err.Number
		Exit Function
	Else
		MoveComputer = 0
	End If
	On Error GoTo 0
	WScript.Sleep(3000)
	Set objComputer = GetObject("LDAP://CN="&strComputer&","&strMoveTo)
	If intType = 1 Then
		intLoc = InStr(1,strObject, ",", 1)
		strOU = Mid(strObject, intLoc+1)
		objComputer.desktopProfile = strOU
	End If
	If IntType = 2 Then
		objComputer.desktopProfile = vbNull
	End If
	objComputer.Setinfo
End Function

'Get info from AD about computer objects
Sub GetComputerObjectInfo()
	rsData.Fields.Append "Computer",adVarChar,MaxCharacters
	rsData.Fields.Append "DN",adVarChar,MaxCharacters
	rsData.Fields.Append "LastLogon",adDBTimeStamp,MaxCharacters
	rsData.Fields.Append "DaysOld",adDouble,MaxCharacters
	rsData.Fields.Append "WhenCreated",adDBTimeStamp,MaxCharacters
	rsData.Fields.Append "OS",adVarChar,MaxCharacters
	rsData.Fields.Append "Description",adVarChar,MaxCharacters
	rsData.Open
	Set objConn = CreateObject("ADODB.Connection")
	objConn.Open "Provider=ADsDSOObject"
	Set objCommand =CreateObject("ADODB.Command")
	objCommand.ActiveConnection = objConn
	objCommand.Properties("Page Size") = 99999
	objCommand.Properties("Chase referrals") = &H60
	objCommand.CommandText = "<LDAP://"&strDefaultNamingContext&">;(objectCategory=computer);distinguishedName,whenCreated,lastLogonTimeStamp,cn,operatingSystem,description;subtree"
	Set objRecordSet = objCommand.Execute
	Do Until objRecordSet.EOF
		blnGo = True
		'exclude quarantine OUs
		If InStr(1,objRecordSet.Fields("distinguishedName").value, GblstrComputerQuarantineDisabled&","&strDefaultNamingContext, 1) <> 0 Then
			blnGo = False
		End If
		If InStr(1,objRecordSet.Fields("distinguishedName").value, GblstrComputerQuarantineActive&","&strDefaultNamingContext, 1) <> 0 Then
			blnGo = False
		End If
		'search excluded computer names list and exclude if found
		For Each strWorkName In arrExcludedComputers
			If UCase(objRecordSet.Fields("cn").value) = ucase(trim(strWorkName)) Then
				blnGo = False
			End If
		Next
		If blnGo = True Then
			'search excluded OU names and exclude if object matches OU name
			For Each strOUName In arrExcludedOUs
				If strOUName <> "" Then
					intLenName = len(strOUName)
					intLoc = Len(objRecordSet.Fields("cn").value) + 5
					If UCase(strOUName) = ucase(Mid(objRecordSet.Fields("distinguishedName").value,intLoc,intLenName)) Then
						blnGo = False
					End If
				End If
			Next
		End If
		If blnGo = True Then
			If blnExcludeServers = True Then
				blnOSMatch = False
				For Each strOS In arrWorkstations
					If Instr(1,objRecordSet.Fields("operatingSystem").value,strOS,1) > 0 Then
						blnOSMatch = True
					End If
				Next
				If blnOSMatch = True Then
					blnGo = True
				Else
					blnGo = False
				End if
			Else 
				blnGo = True
			End If
		End If
		If blnGo = True Then 
			gblMoveandDisable = True
			rsData.AddNew
			rsData("Computer") = objRecordSet.Fields("cn").value
			rsData("WhenCreated") = objRecordSet.Fields("whenCreated").value
			rsData("DN") = objRecordSet.Fields("distinguishedName").value
			If IsNull(objRecordSet.Fields("operatingSystem").value) = False Then
				strConverted = RemoveUnicode(objRecordSet.Fields("operatingSystem").value)
				rsData("OS") = strConverted
			End If
			On Error Resume Next
			Set objLastLogon = objRecordSet.Fields("lastLogonTimeStamp").value
			If Err.Number = 0 Then
				dtLastLogon = ConvertTime(objLastLogon)
			Else
				dtLastLogon = #1/1/1601#
			End If
			On Error GoTo 0
			rsData("LastLogon") = dtLastLogon
			rsData("DaysOld") = DateDiff("d",dtLastLogon, Now)
			'cleanup descriptions
			arrDesc = objRecordSet.Fields("description").value
			If VarType(arrDesc) <> 8204 Then
				strDescription = ""
			Else
				strDescription = arrDesc(0)
			End If
			rsData("Description") = strDescription
			rsData.Update
			iStart = InStr(1,strDescription, "::Account Automatically ", 1) 
			If iStart > 0 Then
				Call SetDescription(objRecordSet.Fields("distinguishedName").value,4)
				If InStr(1, objRecordSet.Fields("operatingSystem").value, "Server", 1) > 0 Then
					GblstrServerEmailMessage = GblstrServerEmailMessage & "Cleaned up description field for "&objRecordset.Fields("cn").value&"."&vbCrLf
				Else
					GblstrWorkstationEmailMessage = GblstrWorkstationEmailMessage & "Cleaned up description field for "&objRecordset.Fields("cn").value&"."&vbCrLf
				End If
			End If	
		End If
		objRecordSet.MoveNext
	Loop
End Sub

'Converts AD time to usable time
Function ConvertTime(adtime)
	Set objDate = adtime
	lngHigh = objDate.HighPart
	lngLow = objDate.LowPart
	If lngLow<0 Then
		lngHigh = lngHigh + 1
	End If 
	If lngHigh=0 And lngLow=0 Then
		dtmDate = #1/1/1601#
	Else
		dtmDate = #1/1/1601# + (((lngHigh*(2^32))+lngLow)/600000000-lngTimeBias)/1440
	End If
	ConvertTime = dtmDate
End Function

'Determines the Time Zone
Function GetTimeZoneBias()
	HKLM = &H80000002
	Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
	strKeyPath = "System\CurrentControlSet\Control\TimeZoneInformation"
	strValueName = "ActiveTimeBias"
	objReg.GetDWORDValue HKLM,strKeyPath,strValueName,dwValue
	GetTimeZoneBias = dwValue
End Function

'Convert a text file into an Array
'If file does not exists returns an empty array
Function ConvertTextToArray(strTextFile)
 	intCount = 0
 	ReDim arrTemp(intCount)
 	If filesys.FileExists(strTextFile) = True Then
		Set objFile = filesys.OpenTextFile(strTextFile, 1)
		Do While objFile.AtEndOfStream = False
			strLine = Trim(objFile.ReadLine)
			If strLine <> "" And Left(strLine,1) <> ";" Then
				ReDim Preserve arrTemp(intCount)
				arrTemp(intCount) = strLine
				intCount = intCount + 1				
			End if
		Loop
	End If
	ConvertTextToArray = arrTemp
End Function

'Emails message to specified address
Sub Email(strSub, strMessage)
	Set objEmail = CreateObject("CDO.Message")
    objEMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
 	objEMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strSMTP
 	objEMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
 	objEMail.Configuration.Fields.Update 
    objEmail.From = strFrom
    objEmail.To = strTo
'    objEmail.BCC = "recipientsbackup@mydomain.local"
    objEmail.Subject = strSub
    objEmail.Textbody =  strMessage
    objEmail.Send
End Sub

'Get Folder of current executing script
Function ScriptPath()
	ScriptPath = Left(WScript.ScriptFullName, Len(WScript.ScriptFullName) - Len(WScript.ScriptName)-1)
End Function
