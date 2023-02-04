Imports System
Imports System.DirectoryServices
Imports System.IO
Imports System.Runtime.InteropServices

Module Module1
    Dim DirPath As String = Nothing
    Dim ADGroup As String = Nothing '"CH2MHILL ITSSG Notifications"


    Sub Main()

        'Setup run 
        Console.WriteLine()
        Console.WriteLine("Please enter the AD group you want to get membership for: ")
        ADGroup = Console.ReadLine()

        Do Until ADGroup IsNot Nothing
            Console.WriteLine("Please enter something")
        Loop

        Console.WriteLine()
        Console.WriteLine("File will be generated where this program is run.")

        'used for testing so i get the logs quicker
        If My.Application.Info.DirectoryPath.Contains("C:\Users\sfreema4\Documents\Visual Studio 2012\Projects") Then
            DirPath = "C:\Users\sfreema4\Desktop"
        Else
            DirPath = My.Application.Info.DirectoryPath
        End If
        'setup log with headers
        Writelog("SetColumns")
        'Get ADgroup desired.
        ListADGroupMembers(ADGroup)

    End Sub

    Public Sub ListADGroupMembers(ByVal GN As String)

        Dim GroupName As String = GN '"G_All_IT_Users"
        Dim GroupMembers As System.Collections.Specialized.StringCollection = GetGroupMembers("CH2MHILL", GN) ' sets up string collection and uses the function to get the group memberships
    End Sub

   

    Public Function GetGroupMembers(ByVal strDomain As String, ByVal strGroup As String) As System.Collections.Specialized.StringCollection

        'Connect to AD
        Dim GroupMembers As New System.Collections.Specialized.StringCollection()
        Dim DirectoryRoot As New DirectoryEntry("LDAP://" & strDomain)
        Dim DirectorySearch As New DirectorySearcher(DirectoryRoot, "(CN=" & strGroup & ")")
        Dim DirectorySearchCollection As SearchResultCollection = DirectorySearch.FindAll()

        'searches each returned group, which should be one but it tends to work better this way
        For Each DirectorySearchResult As SearchResult In DirectorySearchCollection
            Dim ResultPropertyCollection As ResultPropertyCollection = DirectorySearchResult.Properties
            Dim GroupMemberDN As String
            For Each GroupMemberDN In ResultPropertyCollection("member") ' verifies the AD group is legit
                'tries to get the member information.  have in a try block just in case a propertiy is misconfigured
                Try
                    Dim DirectoryMember As New DirectoryEntry("LDAP://" & GroupMemberDN)
                    Dim DirectoryMemberProperties As System.DirectoryServices.PropertyCollection = DirectoryMember.Properties
                    Dim DirectoryItem As Object = DirectoryMemberProperties("sAMAccountName").Value
                    Dim DirectoryEmail As Object = DirectoryMemberProperties("mail").Value
                    Dim DirectoryUPN As Object = DirectoryMemberProperties("userPrincipalName").Value
                    Dim uac As Object = DirectoryMemberProperties("userAccountControl").Value


                    If DirectoryMember.SchemaClassName = "group" Then
                        ' this is a group that was found.
                        ' restart the gathering process with the new group.  as each new group is gathered and finished, code goes back to the previous group to finish that or get the next group
                        ListADGroupMembers(DirectoryMember.Name.Remove(0, 3))

                    End If

                    If DirectoryMember.SchemaClassName = "user" Then
                        ' user found, no more groups to parse
                        ' this is a user.
                        If Nothing IsNot DirectoryItem Then
                            ' check the ad account is enabled
                            If AccEnabled(uac) = 1 Then
                                ' adds the member to the grou member string.  can probably do away with this as i am writing th elog in the next line as it goes rathr than all at once
                                GroupMembers.Add(DirectoryItem.ToString())
                                ' write data to log
                                Writelog(DirectoryItem + ", " + DirectoryEmail + ", " + DirectoryUPN)

                            End If
                        End If
                    End If
                Catch ex As Exception

                End Try
            Next GroupMemberDN
        Next DirectorySearchResult

        Return GroupMembers

    End Function

    ' check account is active or not.

    Function AccEnabled(ByVal uac As String) As String  ' looks to define the account, enabled, locked, etc.

        Dim aret As Integer = 0
        Select Case uac
            Case 512 'Enabled 
                aret = 1
            Case 514 ': ACCOUNTDISABLE()
                aret = 0
            Case 528 ': Enabled(-LOCKOUT)
                aret = 1
            Case 530 ': ACCOUNTDISABLE(-LOCKOUT)
                aret = 0
            Case 544 ': Enabled(-PASSWD_NOTREQD)
                aret = 1
            Case 546 ': ACCOUNTDISABLE(-PASSWD_NOTREQD)
                aret = 0
            Case 560 ': Enabled(-PASSWD_NOTREQD - LOCKOUT)
                aret = 1
            Case 640 ': Enabled(-ENCRYPTED_TEXT_PWD_ALLOWED)
                aret = 1
            Case 2048 ' : INTERDOMAIN_TRUST_ACCOUNT()
                aret = 1
            Case 2080 ': INTERDOMAIN_TRUST_ACCOUNT(-PASSWD_NOTREQD)
                aret = 1
            Case 4096 ': WORKSTATION_TRUST_ACCOUNT()
                aret = 1
            Case 8192 ': SERVER_TRUST_ACCOUNT()
                aret = 1
            Case 66048 ': Enabled(-DONT_EXPIRE_PASSWORD)
                aret = 1
            Case 66050 ': ACCOUNTDISABLE(-DONT_EXPIRE_PASSWORD)
                aret = 0
            Case 66064 ': Enabled(-DONT_EXPIRE_PASSWORD - LOCKOUT)
                aret = 1
            Case 66066 ': ACCOUNTDISABLE(-DONT_EXPIRE_PASSWORD - LOCKOUT)
                aret = 0
            Case 66080 ': Enabled(-DONT_EXPIRE_PASSWORD - PASSWD_NOTREQD)
                aret = 1
            Case 66082 ': ACCOUNTDISABLE(-DONT_EXPIRE_PASSWORD - PASSWD_NOTREQD)
                aret = 0
            Case 66176 ': Enabled(-DONT_EXPIRE_PASSWORD - ENCRYPTED_TEXT_PWD_ALLOWED)
                aret = 1
            Case 131584 ': Enabled(-MNS_LOGON_ACCOUNT)
                aret = 1
            Case 131586 ': ACCOUNTDISABLE(-MNS_LOGON_ACCOUNT)
                aret = 0
            Case 131600 ': Enabled(-MNS_LOGON_ACCOUNT - LOCKOUT)
                aret = 1
            Case 197120 ': Enabled(-MNS_LOGON_ACCOUNT - DONT_EXPIRE_PASSWORD)
                aret = 1
            Case 532480 'SERVER_TRUST_ACCOUNT - TRUSTED_FOR_DELEGATION (Domain Controller) 
                aret = 1
            Case 1049088 ': Enabled(-NOT_DELEGATED)
                aret = 1
            Case 1049090 ': ACCOUNTDISABLE(-NOT_DELEGATED)
                aret = 0
            Case 2097664 ': Enabled(-USE_DES_KEY_ONLY)
                aret = 1
            Case 2687488 ': Enabled(-DONT_EXPIRE_PASSWORD - TRUSTED_FOR_DELEGATION - USE_DES_KEY_ONLY)
                aret = 1
            Case 4194816 ': Enabled(-DONT_REQ_PREAUTH)
                aret = 1
            Case Else
                aret = 0
        End Select

        AccEnabled = aret

    End Function

    Public Sub Writelog(ByVal LogEntry As String)

        Using sw As StreamWriter = File.AppendText(DirPath + "\" + ADGroup.Replace(" ", "").Replace("/", "") + ".csv")
            If LogEntry = "SetColumns" Then
                sw.WriteLine("SamAccountName, Email, UPN")
            Else
                sw.WriteLine(LogEntry)
            End If
            sw.Flush()
            sw.Close()

        End Using

    End Sub

End Module
