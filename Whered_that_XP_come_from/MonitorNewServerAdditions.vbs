' by http://kevinisms.fason.org

'Set these variables to your evironment
strEmails = "helpdesk@my.domail.local;seniorhelpdesk@my.domain.local" 'separate with semi-colon
strSMTPServer = "smtpsmarthost.my.domain.local"
accessdb = ScriptPath & "\Servers.mdb"
strDomainLDAP = "dc=my,dc=domain,dc=local"

intCount = 0
strMessage = ""
arrEmails = Split(strEmails, ";")
Set objConnection = CreateObject("ADODB.Connection")
objConnection.Open "Provider=ADsDSOObject;"
Set objCommand = CreateObject("ADODB.Command")
objCommand.ActiveConnection = objConnection
objCommand.CommandText = "<LDAP://"&strDomainLDAP&">;(objectCategory=Computer);distinguishedName;subtree"  
objCommand.Properties("Page Size")=99999
Set objRecordSet = objCommand.Execute
Do Until objRecordset.EOF
	strDN =objRecordset.Fields("distinguishedName")
	Set objComputer = GetObject("LDAP://"&strDN)
	strOSType = ""
	On Error Resume Next
	strOSType = objComputer.Get("operatingSystem")
	On Error GoTo 0
	If strOSType <> "" Then
		If InStr(1, strOSType, "server", 1) > 0 Then
			strSID = cstr(sid2hexstr(objComputer.Get("objectSID")))
			strServer = objComputer.Get("cn")
			objComputer.GetInfoEx Array("canonicalName"), 0
			strCN = objComputer.Get("canonicalName")
			If IsNew(strSID) = True Then
				intCount = intCount + 1
				InsertNewServer strServer, strSID, strCN 
				strMessage = strMessage & "Server "&strServer&" running "&strOSType&" is new and is found at: " & strCN & VbCrLf &VbCrLf			
			End If
		End If
	End If
	objRecordset.MoveNext
Loop
objConnection.Close

If intCount > 0 Then
	SendEmail
End If
WScript.Quit


Sub SendEmail
	For Each Email In arrEmails
		 Set objMail = CreateObject ("CDO.Message") 
		 objMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
		 objMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strSMTPServer
		 objMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
		 objMail.Configuration.Fields.Update 
		 objMail.From = "NewServersAlert@my.domain.local"
		 objMail.To = Email
		 If intCount = 1 Then
		 	objMail.Subject = "New server has joined the my.domain.local domain"
		 Else 
		 	 	objMail.Subject = intCount&" new servers have joined the my.domain.local domain"
		 End If
		 objMail.TextBody = strMessage
		 objMail.send
	Next
End Sub

Sub InsertNewServer(strNewServer, strNewSID, strNewCN)
	Set Connection = WScript.CreateObject("ADODB.Connection")
	Connection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&accessdb&";" 
	Set cmdInsert = Connection.Execute("INSERT INTO Servers (ServerName, SID, Location) VALUES ('"&strNewServer&"', '"&strNewSID&"', '"&strNewCN&"')",, adCmdTxt)
End Sub

Function IsNew(strSearchSID)
	Set Connection = WScript.CreateObject("ADODB.Connection")
	Connection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&accessdb&";" 
	Set RS = Connection.Execute("SELECT * FROM Servers Where SID ='"&strSearchSID&"'",, adCmdTxt)
	blnFound = False
	While Not RS.EOF
		blnFound = True
	 	RS.MoveNext
  	Wend
  	If blnFound = True Then
  		IsNew = False
  	Else
  		IsNew = True
  	End if
	
End Function


Function sid2hexstr (sid_value)
	tmpstr = ""
	For i = LBound(sid_value) To UBound(sid_value)
	     tmpstr = tmpstr & Hex(AscB(MidB(sid_value,i+1,1)) \ 16) & Hex(AscB(MidB(sid_value,i+1,1)) Mod 16)
	Next
	 sid2hexstr = tmpstr
End Function

Function ScriptPath()
	ScriptPath = Left(WScript.ScriptFullName, _
	   Len(WScript.ScriptFullName) - Len(WScript.ScriptName) - 1)
End Function
