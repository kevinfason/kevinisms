' by http://kevinisms.fason.org
Set env = CreateObject("Microsoft.SMS.TSEnvironment")
Set oTSProgressUI = CreateObject("Microsoft.SMS.TSProgressUI")
oTSProgressUI.CloseProgressDialog()
strTSPassword = env("OSDPASSWORD") 
blnGo = True
While blnGo = True
	strMyPass=Inputbox("Please enter the Password to continue")
	If strTSPassword = strMyPass Then
		blnGo = False
	Else
		MsgBox "Incorrect password entered, please try again", 16, "Wrong Password"
	End If
Wend
WScript.Quit(0)




