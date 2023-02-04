'From http://kevinisms.fason.org
'Edit these lines to your environment'
' 35, 38, 29
' 35 and 38 get the domain suffix you want them to be
' 29 needs the LDAP string to connect to a correct DC
 
strXLS = ScriptPath()&"\prod_List.xlsx"
Set objExcel = CreateObject("EXCEL.APPLICATION")
Set objWorkBook = objExcel.Workbooks.Open(strXLS)
objExcel.Visible = True
Set objSheet = objWorkBook.WorkSheets("prod_List")
For iRow = 2 To 10000
	If objSheet.Cells(iRow,1).Value = "" Then
		Exit For
	End If
	strName = objSheet.Cells(iRow,1).Value
	objSheet.Cells(iRow,8).Value = UpdateUPN(strName)
Next

WScript.Quit(0)
'--------------------------------------------------------------
Function UpdateUPN(sUserName)
	Const ADS_SCOPE_SUBTREE = 2
	sNewUPN = ""
	Set oConnection = CreateObject("ADODB.Connection")
	Set oCommand =   CreateObject("ADODB.Command")
	oConnection.Provider = "ADsDSOObject"
	oConnection.Open "Active Directory Provider"
	Set oCommand.ActiveConnection = oConnection
	oCommand.Properties("Page Size") = 1000
	oCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 
	oCommand.CommandText = "SELECT AdsPath,samAccountName,userPrincipalName FROM 'LDAP://dc=my,dc=domain,dc=com' WHERE samAccountName='"&sUsername&"'"  
	Set oRecordSet = oCommand.Execute
	Do Until oRecordSet.EOF
	    sUser = oRecordSet.Fields("ADsPath").Value
	    sUPN = oRecordSet.Fields("userPrincipalName").Value
		If IsNull(sUPN) = True Then
			sNewUPN = sUserName&"@domain.com"
		Else
			aUPN = Split(sUPN, "@")
			sNewUPN = aUPN(0)&"@domain.com"
		End If
	    Set oUser =  GetObject(sUser)
	    oUser.userPrincipalName = sNewUPN
	    On Error Resume Next
	    oUser.SetInfo
	    If Err.Number <> 0 Then
	    	sNewUPN = "Updating Error:"&Err.Number&" "&Err.Description
	    End If
	    oRecordSet.MoveNext
	Loop
	UpdateUPN = sNewUPN
End Function

Function ScriptPath()
	ScriptPath = Left(WScript.ScriptFullName, Len(WScript.ScriptFullName) - Len(WScript.ScriptName)-1)
End Function
