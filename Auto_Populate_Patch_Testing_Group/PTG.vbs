'-----------------------------------------------------------------
'User defined variables
'-----------------------------------------------------------------
blnTestRun = True 'If True, the computer account and user accounts will not be added to the group.  Emails will not be sent to the user either.
blnEmailUser = False 'do you want to email the end user that their computer has been added to the patch testing group?  Test Run variable has to be set to False otherwise this is skipped.
blnEmailTestUser = False 'If set to True then the test email account will receive the indivdual emails as well
strPTGNameUser = "PTGTestGroupUsers" 'Name of group to contain user accounts
strPTGNameComputer = "PTGTestGroupComputers" 'Name of group to contain user accounts
intQuantity = 1000 'maximum number of computer accounts in the group
intMaxAddLimit = 50 'maximum number of computer accounts to add each time the script is run
intMinBuildDays = 30 'number of days past build date to look at for adding to the group.  Make sure that User Device Affinity has had enough time to generate this
strEmailFile = ScriptPath()&"\EmailToUser.txt" 'path to the standard email template
strSMTPServer = "smarthost.mycompany.com" 'SMTP server to use
strEmailsTech = "helpdesk@mycompany.com" ' separate with semi-colon and space.   Ex:  "me@company.com; you@company.com"
strEmailFromTech = "PTGScript@mycompany.com" 'From address that goes to email addresses that receive test and summary messages
strEmailFromUser = "ITCommunicationsmycompany.com" 'From address that users will see
strEmailTest = "helpdesk@mycompany.com" 'Test email account that will receive the individual emails
strEmailSubjectUser = "Your computer has been added to the patch testing group" 'Subject in email that goes out to users
strSQLServer = "SCCMDBserver.mycompany.com" 'FQDN of SQL server hosting ConfigMgr database
strDB = "CM_DB" 'DB name of ConfigMgr
'-----------------------------------------------------------------
'End of user defined variables
'-----------------------------------------------------------------



'Get current count of computer accounts in the PTG-Computer
intNumberOfComputers = GetNumberOfAccounts(strPTGNameComputer)
'If there are enough computers in the group, exit and send email
If intNumberofComputers >= intQuantity Then
	Call Email("Patch Testing Group Script Ran", strEmailsTech, strEmailFromTech, "There are currently "&intNumberofComputers&" computers in the "&strPTGNameComputer&" group.  No need to add additional computers.")
	WScript.Quit(0)
End If
'create recordset to contain the potential computer accounts to add
Const adVarChar = 200
Const MaxCharacters = 255
Set rsList = CreateObject("ADOR.Recordset")
rsList.Fields.Append "Name", adVarChar, MaxCharacters
rsList.Fields.Append "InstallDate", adVarChar, MaxCharacters
rsList.Fields.Append "UserName", adVarChar, MaxCharacters
rsList.Fields.Append "Email", adVarChar, MaxCharacters
rsList.Fields.Append "DN", adVarChar, MaxCharacters
rsList.Fields.Append "Added", adVarChar, MaxCharacters
rsList.Open
'Query SQL and populate the recordset with computers
Call PopulateList()
'start adding potential list of computers and users to the groups
intAdded = 0
Set objGroupList = CreateObject("Scripting.Dictionary")
Call AddToGroups()
'send summary to specified email addresses
Call SendSummary()
'end of script
WScript.Quit(0)


'----------------------------------------------------------------
Sub SendSummary()
	sMessage = ""
	iCount = 0
	rsList.MoveFirst
	Do While Not rsList.EOF
		If rsList.Fields.Item("Added").Value = "TRUE" Then
			sPC = rsList.Fields.Item("Name").Value
			If Len(sPC) < 10 Then
				sPC = sPC & vbTab
			End If
			sUser = rsList.Fields.Item("UserName").Value
			If Len(sUser) < 15 Then
				sUser = sUser & vbTab
			End If
			sMessage = sMessage & "Computer: "&sPC&"  "&vbTab&"Install Date: "&rsList.Fields.Item("InstallDate").Value&vbTab&"User:  "&sUser&" "&vbTab&"Email: "&rsList.Fields.Item("Email").Value & vbCrLf
			iCount = iCount + 1
		End If
		rsList.MoveNext
	Loop
	If iCount > 0 Then
		sMessage = intAdded&" computers were added to the patch testing group:"&vbCrLf&vbCrLf&sMessage
	Else
		sMessage = "No computers were added to the patch testing group.  There are currently "&intNumberOfComputers&" computers in the group."
	End If
	If blnTestRun = True Then
		sMessage = "This is a test run only.  Group membership has not been modified and no emails sent to users."&vbCrLf&vbCrLf&sMessage
	End If	
	Call Email("Patch testing group script summary", strEmailsTech, strEmailFromTech, sMessage)
End Sub

Sub AddToGroups()
	sCGroupDN = GetDN(strPTGNameComputer,1)
	sUGroupDN = GetDN(strPTGNameUser,1)
	If sUGroupDN = "" Then
		Call Email("Patch Testing Group Script Ran", strEmailsTech, strEmailFromTech, "Unable to find group "&strPTGNameUser&" in the domain.  Script is unable to run.")
		WScript.Quit(53)
	End if
	rsList.MoveFirst
	Do While Not rsList.EOF
		If intAdded < intMaxAddLimit Then
			sCAccountDN = GetDN(rsList.Fields.Item("Name").Value, 2)
			sUAccountDN = GetDN(FormatUsername(rsList.Fields.Item("UserName").Value), 3)
			'check if the computer account is a member of the computer PTG 
			If IsMember(sCAccountDN, sCGroupDN) = False Then
				If blnTestRun = False Then
					If AddToGroup(sCAccountDN, sCGroupDN) = True Then
						intAdded = intAdded + 1
						strMessage = FormatMessage(rsList.Fields.Item("Name").Value)
						rsList("Added") = "TRUE"
						rsList.Update
						If blnEmailUser = True Then
							Call Email(strEmailSubjectUser, rsList.Fields.Item("Email").Value, strEmailFromUser, strMessage)
						End If
						If blnEmailTestUser = True Then
							Call Email(strEmailSubjectUser, strEmailTest, strEmailFromUser, strMessage)
						End If
						'now add the user account to the user PTG
						If IsMember(sUAccountDN, sUGroupDN) = False Then
							Call AddToGroup(sUAccountDN, sUGroupDN)
						End If
					End If
				Else
					If blnEmailTestUser = True Then
						strMessage = FormatMessage(rsList.Fields.Item("Name").Value)
						Call Email(strEmailSubjectUser, strEmailTest, strEmailFromUser, strMessage)
					End If
					rsList("Added") = "TRUE"
					rsList.Update
					intAdded = intAdded + 1									
				End If	
			Else
				If IsMember(sUAccountDN, sUGroupDN) = False Then
					If blnTestRun = False Then
						Call AddToGroup(sUAccountDN, sUGroupDN)
					End If
				End If
			End If
		End If
		rsList.MoveNext
	Loop
End Sub

Function FormatMessage(sComputer)
	sMessage = ""
	Set filesys = CreateObject("Scripting.FileSystemObject")
	Set oFile = filesys.OpenTextFile(strEmailFile, 1)
	Do While oFile.AtEndOfStream = False
		sLine = Trim(oFile.ReadLine)
		If InStr(1, sLine, "%COMPUTER%", 1) > 0 Then
			sLine = Replace(sLine, "%COMPUTER%", sComputer)
		End If
		sMessage = sMessage & sLine & vbCrLf
	Loop
	FormatMessage = sMessage
End Function

Function AddToGroup(sAccountDN, sGroupDN)
	On Error Resume Next
	Set oGroup = GetObject("LDAP://"&sGroupDN)
	oGroup.Add("LDAP://"&sAccountDN)
	If Err.Number = 0 Then
		AddToGroup = True
	Else
		AddToGroup = False
	End If
	On Error GoTo 0
End Function

Function ScriptPath()
	ScriptPath = Left(WScript.ScriptFullName, Len(WScript.ScriptFullName) - Len(WScript.ScriptName)-1)
End Function

Function IsMember(sAccountDN, sGroupDN)
	IsMember = False
	Set oGroup = GetObject("LDAP://"&sGroupDN)
	For Each oMember In oGroup.Members
		If ucase(sAccountDN) = ucase(oMember.distinguishedName) Then
			IsMember = True
			Exit For
		End if
	Next
End Function

Function FormatUserName(sUser)
	iPos = InStr(1, sUser, "\", 1)
	FormatUserName = Mid(sUser, iPos+1)
End Function

Sub PopulateList()
	Set oConn = CreateObject("ADODB.Connection")
	oConn.Open "Driver={SQL Server};Server="&strSQLServer&";Database="&strDB&";Trusted_Connection=yes"
	sQuery = "select TOP "&intQuantity&" sys.name0, OS.InstallDate0, RUSER.Unique_User_Name0, RUSER.Mail0, SYS.Distinguished_Name0 " _
				& "FROM v_R_System SYS " _
				& "outer apply (select top 1 * from v_UserMachineRelation UMR1 where UMR1.MachineResourceID = sys.ResourceID order by rowversion DESC) as UMR " _
				& "JOIN v_GS_OPERATING_SYSTEM OS on SYS.ResourceID=OS.ResourceID " _
				& "JOIN v_R_User RUSER on UMR.UniqueUserName=RUSER.Unique_User_Name0 " _
				& "WHERE SYS.Operating_System_Name_and0 LIKE '%Workstation%' AND DATEDIFF(DAY,OS.InstallDate0,GETDATE()) > "&intMinBuildDays&" AND RUSER.Mail0 IS NOT NULL AND SYS.Distinguished_Name0 IS NOT NULL " _
				& "ORDER BY OS.InstallDate0 DESC"
	Set oRst = oConn.Execute(sQuery,, adCmdTxt)
	While Not oRst.Eof
		rsList.AddNew
		rsList("Name") = oRst.Fields("Name0").Value
		rsList("InstallDate") = oRst.Fields("InstallDate0").Value
		rsList("UserName") = oRst.Fields("Unique_User_Name0").Value
		rsList("Email") = oRst.Fields("Mail0").Value
		rsList("DN") = oRst.Fields("Distinguished_Name0").Value
		rsList("Added") = "True"
		rsList.Update
		oRst.MoveNext
	Wend
End Sub

Function GetNumberOfAccounts(sGroup)
	iCount = 0
	sDN = GetDN(sGroup, 1)
	If sDN="" Then
		'quit if unable to find group
		Call Email("Patch Testing Group Script Ran", strEmailsTech, strEmailFromTech, "Unable to find group "&sGroup&" in the domain.  Script is unable to run.")
		WScript.Quit(53)
	End if
	Set oGroup = GetObject("LDAP://"&sDN)
	oGroup.GetInfo
	aMemberOf = oGroup.GetEx("member")
	For Each sComputer In aMemberOf
		iCount = iCount + 1
	Next
	GetNumberOfAccounts = iCount	
End Function

Function GetDN(sTypeName,intType)
	Select Case intType
		Case 1
			sCat = "Group"
			sAttr = "name"
		Case 2
			sCat = "Computer"
			sAttr = "name"
		Case 3
			sCat = "User"
			sAttr = "sAMAccountName"
	End Select
	sDN = ""
	Set oRootDSE = GetObject("LDAP://RootDSE")
	sDefaultNamingContext = oRootDSE.Get("defaultNamingContext")
	Set oConnection = CreateObject("ADODB.Connection")
	oConnection.Open "Provider=ADsDSOObject;"
	Set oCommand = CreateObject("ADODB.Command")
	oCommand.ActiveConnection = oConnection
	oCommand.CommandText = "<LDAP://"&sDefaultNamingContext&">;(&(objectCategory="&sCat&")("&sAttr&"="&sTypeName&"));distinguishedName;subtree"  
	Set oRecordSet = oCommand.Execute
	sDN = ""
	Do Until oRecordSet.EOF
		sDN = oRecordSet.Fields("distinguishedName").Value
		oRecordSet.MoveNext
	Loop
	GetDN = sDN
End Function

Sub Email(sSubject, sEmailTo, sEmailFrom, sBody)
	Set oEmail = CreateObject("CDO.Message")
    oEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
 	oEMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strSMTPServer
 	oEMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
 	oEMail.Configuration.Fields.Update 
    oEmail.From = sEmailFrom
   	oEmail.To = sEmailTo
    oEmail.Subject = sSubject
    oEmail.Textbody = sBody
    oEmail.Send
    Set oEmail = Nothing
End Sub

