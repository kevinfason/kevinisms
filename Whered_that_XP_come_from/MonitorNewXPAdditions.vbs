' by http://kevinisms.fason.org

'Set these variables to your evironment
strEmails = "helpdesk@my.domail.local;seniorhelpdesk@my.domain.local"  'separate with semi-colon
strSMTPServer = "smtpsmarthost.my.domain.local"
accessdb = ScriptPath & "\XP.mdb"
strDomainLDAP = "dc=my,dc=domain,dc=local"
strDomainFQDN = "my.domain.local"
strSMTPDomain = "my.domain.local" 'If email from domain different from domain being monitored

intCount = 0
strMessage = ""
arrEmails = Split(strEmails, ";")
Set objConn = CreateObject("ADODB.Connection")
objConn.Provider = "ADsDSOObject"
objConn.Open "ADs Provider"
Set objComm = CreateObject("ADODB.Command")
objComm.ActiveConnection = objConn
objComm.CommandText = "<LDAP://"&strDomainLDAP&">;(&(objectCategory=Computer)(operatingSystem=*XP*));distinguishedName;subtree"
objComm.Properties("Page Size")=99999
Set objRS = objComm.Execute
Do Until objRS.EOF
	strDN =objRS.Fields("distinguishedName")
	Set objComputer = GetObject("LDAP://"&strDN)
	strOSType = objComputer.Get("operatingSystem")
	strSID = cstr(sid2hexstr(objComputer.Get("objectSID")))
	strComputer = objComputer.Get("cn")
	objComputer.GetInfoEx Array("canonicalName"), 0
	strCN = objComputer.Get("canonicalName")
	If IsNew(strSID) = True Then
		intCount = intCount + 1
		InsertNewServer strComputer, strSID, strCN
		strDNSIP = LookupDNS(strComputer) 
		strADSite = GetADSite(strDNSIP)
		strMessage = strMessage & "Computer "&strComputer&" running "&strOSType&" is new and is found at: " & strCN&VbCrLf&"   IP Address:"&strDNSIP&"   AD Site:"&strADSite&vbcrlf&VbCrLf			
	End If
	objRS.MoveNext
Loop
'objConnection.Close

If intCount > 0 Then
	SendEmail
End If
WScript.Quit
'--------------------------------------------------------------
'Determines the AD site of the computer name that is passed to the function.  
'Returns AD Site name if found, otherwise returns blank string
Function GetADSite(strIPA)
	GetADSite = ""
	strADSite = LookUpADSite(strIPA)
	If strADSite <> "" Then
		GetADSite = strADSite
	Else
	End If		
End Function

Function CalcSubnet(strAddress, strMask)
	intSubnetLength = SubnetLength(strMask)
	CalcSubnet = BinaryToString(Left(StringToBinary(strAddress), intSubnetLength) & String(32 - intSubnetLength, "0"))
End Function

Function SubnetLength(strMask)
	strMaskBinary = StringToBinary(strMask)
	SubnetLength = Len(Left(strMaskBinary, InStr(strMaskBinary, "0") - 1))
End Function

Function BinaryToString(strBinary)
	For intOctetPos = 1 To 4
		strOctetBinary = Right(Left(strBinary, intOctetPos * 8), 8)
		intOctet = 0
		intValue = 1
		For intBinaryPos = 1 To Len(strOctetBinary)
			If Left(Right(strOctetBinary, intBinaryPos), 1) = "1" Then intOctet = intOctet + intValue
			intValue = intValue * 2
		Next
		If BinaryToString = Empty Then BinaryToString = CStr(intOctet) Else BinaryToString = BinaryToString & "." & CStr(intOctet)
	Next
End Function

Function StringToBinary(strAddress)
	objAddress = Split(strAddress, ".", -1)
	For Each strOctet In objAddress
		intOctet = CInt(strOctet)
		strOctetBinary = ""
		For x = 1 To 8
			If intOctet Mod 2 > 0 Then
				strOctetBinary = "1" & strOctetBinary
			Else
				strOctetBinary = "0" & strOctetBinary
			End If
			intOctet = Int(intOctet / 2)
		Next
		StringToBinary = StringToBinary & strOctetBinary
	Next
End Function

Function LookUpADSite(strIP)
	SubNetCounter = 0
	Set objRootDSE = GetObject("LDAP://RootDSE")
	strConfigurationNC = objRootDSE.Get("configurationNamingContext")
	Set objConnection = CreateObject("ADODB.Connection")
	objConnection.Open "Provider=ADsDSOObject;"
	Set objCommand = CreateObject("ADODB.Command")
	objCommand.ActiveConnection = objConnection
	objCommand.CommandText = "<LDAP://cn=Subnets,cn=Sites," & strConfigurationNC&">;;distinguishedName;subtree"  
	objCommand.Properties("Page Size")=99999
	objCommand.Properties("Chase referrals")=&H60
	Set objRecordSet = objCommand.Execute
	Do Until objRecordset.EOF
		If objRecordset.Fields("distinguishedName").Value <> "CN=Subnets,CN=Sites,"&strConfigurationNC Then
			ReDim Preserve arrADSubnets(SubNetCounter)
			arrADSubnets(SubNetCounter) = objRecordset.Fields("distinguishedName").Value
			SubNetCounter = SubNetCounter + 1
		End If
		objRecordSet.movenext
	Loop
	objRecordSet.close
	For i = 0 To UBound(arrADSubnets)
		CurrentSubnet = Split(arrADSubnets(i),",")
		arrParsedSubnet = Split(CurrentSubnet(0),"=")
		arrADSubnets(i) = arrParsedSubnet(1)
	Next
	For t = 0 To UBound(arrADSubnets)
		arrSubnet1 = Split(arrADSubnets(t),"/")
		Subnet = arrSubnet1(0)
		mask = arrSubnet1(1)
		Select Case Mask
		Case 4
			NetMask = "240.0.0.0"
		Case 5
			NetMask = "248.0.0.0"
		Case 6
			NetMask = "252.0.0.0"
		Case 7
			NetMask = "254.0.0.0"
		Case 8
			NetMask = "255.0.0.0"
		Case 9
			NetMask = "255.128.0.0"
		Case 10
			NetMask = "255.192.0.0"
		Case 11
			NetMask = "255.224.0.0"
		Case 12
			NetMask = "255.240.0.0"
		Case 13
			NetMask = "255.248.0.0"
		Case 14
			NetMask = "255.252.0.0"
		Case 15
			NetMask = "255.254.0.0"
		Case 16
			NetMask = "255.255.0.0"
		Case 17
			NetMask = "255.255.128.0"
		Case 18
			NetMask = "255.255.192.0"
		Case 19
			NetMask = "255.255.224.0"
		Case 20
			NetMask = "255.255.240.0"
		Case 21
			NetMask = "255.255.248.0"
		Case 22
			NetMask = "255.255.252.0"
		Case 23
			NetMask = "255.255.254.0"
		Case 24
			NetMask = "255.255.255.0"
		Case 25
			NetMask = "255.255.255.128"
		Case 26
			NetMask = "255.255.255.192"
		Case 27
			NetMask = "255.255.255.224"
		Case 28
			NetMask = "255.255.255.240"
		Case 29
			NetMask = "255.255.255.248"
		Case 30
			NetMask = "255.255.255.252"
		Case Else
			NetMask = 0
		End Select
		If CalcSubnet(strIP,NetMask) = Subnet Then
			strADSubnet = arrADSubnets(t)
			Exit For
		End If
	Next
	objCommand.CommandText = "<LDAP://cn=Sites," & strConfigurationNC&">;(siteObjectBL=CN="&strADSubnet&",cn=Subnets,cn=Sites,"&strConfigurationNC&");distinguishedName;subtree"  
	Set objRecordSet = objCommand.Execute
	Do Until objRecordset.EOF
		arrADSite=Split(objRecordset.Fields("distinguishedName").Value,",")
		arrADSite1 = Split(arrADSite(0),"=")
		ADSite = arrADSite1(1)
	 objRecordset.movenext
	'Wend
	Loop
	objRecordset.close
	LookUpADSite = ADSite
End Function

'Lookup DNS record IP address
Function LookupDNS(strAHost)
	LookupDNS = ""
	If Len(strAHost) > 16 Then
		Exit Function
	End If
	Set objPingResults = GetObject("winmgmts:{impersonationLevel=impersonate}//./root/cimv2").ExecQuery("Select * from Win32_PingStatus WHERE Address="&Chr(34)&strAHost&Chr(34))
	For Each PingResult In objPingResults
		LookupDNS = PingResult.ProtocolAddress
	Next
End Function

Sub SendEmail
	For Each Email In arrEmails
		 Set objMail = CreateObject ("CDO.Message") 
		 objMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
		 objMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strSMTPServer
		 objMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
		 objMail.Configuration.Fields.Update 
		 objMail.From = "NewXPAlert@"&strSMTPDomain&""
		 objMail.To = Email
		 If intCount = 1 Then
		 	objMail.Subject = "New XP computer has joined the "&strDomainFQDN&" domain"
		 Else 
		 	 	objMail.Subject = intCount&" new XP computers have joined the "&strDomainFQDN&" domain"
		 End If
		 objMail.TextBody = strMessage
		 objMail.send
	Next
End Sub

Sub InsertNewServer(strNewComputer, strNewSID, strNewCN)
	Set Connection = WScript.CreateObject("ADODB.Connection")
	Connection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&accessdb&";" 
	Set cmdInsert = Connection.Execute("INSERT INTO Computers (ServerName, SID, Location) VALUES ('"&strNewComputer&"', '"&strNewSID&"', '"&strNewCN&"')",, adCmdTxt)
End Sub

Function IsNew(strSearchSID)
	Set Connection = WScript.CreateObject("ADODB.Connection")
	Connection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&accessdb&";" 
	Set RS = Connection.Execute("SELECT * FROM Computers Where SID ='"&strSearchSID&"'",, adCmdTxt)
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
