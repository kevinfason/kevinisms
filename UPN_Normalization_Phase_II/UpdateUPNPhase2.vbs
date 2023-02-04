' from http://kevinisms.fason.org
Const MaxToUpdate  = 5000
arrSkippedOUs = Array("OU=ServiceAccounts,OU=Accounts")
Const ADS_SCOPE_SUBTREE = 2
Set oConnection = CreateObject("ADODB.Connection")
Set oCommand =   CreateObject("ADODB.Command")
oConnection.Provider = "ADsDSOObject"
oConnection.Open "Active Directory Provider"
Set oCommand.ActiveConnection = oConnection
oCommand.Properties("Page Size") = 99999
oCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 
oCommand.CommandText = "<LDAP://dc=amr,dc=ch2m,dc=com>;(&(objectCategory=User)(employeeID=*));AdsPath,mail,userPrincipalName;Subtree"

Set oRecordSet = oCommand.Execute
On Error Resume Next
i = 0
Do Until oRecordSet.EOF
	
	If i > MaxToUpdate Then
		WScript.Quit()
	End If
    strUser = oRecordSet.Fields("ADsPath").Value
    For Each strSkippedOU In arrSkippedOUs
    	If InStr(1, strSkippedOU, strUser, 1) = 0 Then
		    strEmail = oRecordSet.Fields("mail").Value
		    strUPN = oRecordSet.Fields("userPrincipalName").Value
		    If InStr(1, strEmail, "@", 1) > 0 Then
				arrEmail = Split(strEmail, "@")
				If UCase(strUPN) <> UCase(arrEmail(0)&"@MYDOMAIN.COM")  AND Instr(arrEmail(0), "INVALID_") = 0 Then
					strNewUPN = arrEmail(0)&"@mydomain.com"
					WScript.Echo "Updating "& strUPN & " to " & strNewUPN
					status("Updating "& strUPN & " to " & strNewUPN)
					Set objUser =  GetObject(strUser)
					objUser.userPrincipalName = strNewUPN
            'This next line pulls the trigger
			   		objUser.SetInfo
				    If Err.Number <> 0 Then
				    	wscript.echo "Updating Error:"&Err.Number&" "&Err.Description
				    	status("Updating Error:" & Err.Number & " " & Err.Description)
				    Else
				    	WScript.Echo "   Updated"				    	
				    End If
				    i = i + 1
				End If		
		    End If
		End If
	Next
    oRecordSet.MoveNext
Loop	
'***************************************************************************************************'
Sub Status(strMessage)
	Dim ts 'As Scripting.TextStream
	Dim fs 'As Scripting.FileSystemObject
	Const ForAppending = 8 'Scripting.IOMode

	
		Set fs = CreateObject("Scripting.FileSystemObject")
		Set ts = fs.OpenTextFile(Wscript.ScriptFullName & ".log", ForAppending, True)
		ts.WriteLine strMessage
		ts.Close
	
  	''''''''''Clean up
		Set ts = Nothing
		Set fs = Nothing
	
End Sub