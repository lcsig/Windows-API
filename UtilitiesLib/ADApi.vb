Imports System.DirectoryServices
Imports System.Diagnostics
Imports System.Security
Imports System.IO
Imports System.Text
Imports System.Runtime.InteropServices

Public Class ADApi
    Private Const UF_ALL As Integer = &H0 '0   All account types. 
    Private Const UF_SCRIPT As Integer = &H1 '1   The logon script executed. This value must be set for LAN Manager 2.0 or Windows NT. 
    Private Const UF_ACCOUNTDISABLE As Integer = &H2 '2   The user's account is disabled. 
    Private Const UF_HOMEDIR_REQUIRED As Integer = &H8 '8   The home directory is required. This value is ignored in Windows NT. 
    Private Const UF_LOCKOUT As Integer = &H10 '16   The account is currently locked out. This value can be cleared to unlock a previously locked account. This value cannot be used to lock a previously unlocked account. 
    Private Const UF_PASSWD_NOTREQD As Integer = &H20 '32   No password is required. 
    Private Const UF_PASSWD_CANT_CHANGE As Integer = &H40 '64   The user cannot change the password. 
    Private Const UF_TEMP_DUPLICATE_ACCOUNT As Integer = &H100 '256   This is an account for users whose primary account is in another domain. This account provides user access to this domain, but not to any domain that trusts this domain. The User Manager refers to this account type as a local user account. 
    Private Const UF_NORMAL_ACCOUNT As Integer = &H200 '512   This is a default account type that represents a typical user. 
    Private Const UF_INTERDOMAIN_TRUST_ACCOUNT As Integer = &H800 '2 048   This is a permit to trust account for a Windows NT domain that trusts other domains. 
    Private Const UF_WORKSTATION_TRUST_ACCOUNT As Integer = &H1000 '4 096   This is a computer account for a Windows NT Workstation or Windows NT Server that is a member of this domain. 
    Private Const UF_SERVER_TRUST_ACCOUNT As Integer = &H2000 '8 192   This is a computer account for a Windows NT Backup Domain Controller that is a member of this domain. 
    Private Const UF_MACHINE_ACCOUNT_MASK As Integer = &H3800 '14 336    
    Private Const UF_ACCOUNT_TYPE_MASK As Integer = &H3B00 '15 104    
    Private Const UF_DONT_EXPIRE_PASSWD As Integer = &H10000 '65 536    
    Private Const UF_MNS_LOGON_ACCOUNT As Integer = &H20000 '131 072    
    Private Const UF_SETTABLE_BITS As Integer = &H33B7B '211 835 

    Public Shared Function List() As DirectoryEntries
        Dim root As DirectoryEntry
        root = New DirectoryEntry("WinNT://" + Environment.MachineName + ",computer")
        Return root.Children
    End Function

    Public Shared Sub DeleteUser(ByVal user As DirectoryEntry)
        List().Remove(user)
    End Sub

    Public Shared Function GetUser(ByVal login As String) As DirectoryEntry
        Dim domain As String = Nothing
        Dim name As String = Nothing

        Utilities.ParseUsername(login, domain, name)

        Dim user As DirectoryEntry = Nothing

        Try
            user = List().Find(name, "user")
        Catch cex As COMException
            If cex.ErrorCode = -2147022675 Then
                'user doesn't exist
                Return Nothing
            Else
                Throw cex
            End If
        End Try
        Return user
    End Function

    Public Shared Function CreateUser(ByVal login As String, ByVal password As String, ByVal groups() As String) As DirectoryEntry
        Dim entries As DirectoryEntries = ADApi.List()

        Dim domain As String = Nothing
        Dim username As String = Nothing
        Utilities.ParseUsername(login, domain, username)

        Logger.WriteEntry("Creating user: " + login)

        Dim newUser As DirectoryEntry = entries.Add(username, "user")
        newUser.Properties("FullName").Add(username)
        newUser.Properties("PasswordExpired").Add(0)
        newUser.Properties("UserFlags").Add(UF_DONT_EXPIRE_PASSWD)
        'newUser.Properties("PasswordAge").Add(0)

        newUser.Invoke("SetPassword", password)

        Logger.WriteEntry("I am: " + System.Security.Principal.WindowsIdentity.GetCurrent().Name)
        Logger.WriteEntry("Saving new user: " + login)

        newUser.CommitChanges()

        Logger.WriteEntry("success")

        Dim grp As DirectoryEntry
        For Each s As String In groups
            grp = entries.Find(s, "group")
            If grp IsNot Nothing Then
                Debug.Print("Adding user to group: " + s)
                grp.Invoke("Add", New Object() {newUser.Path.ToString()})
            End If
        Next

        Return newUser
    End Function

    Public Shared Function AccountIsDisabled(ByVal entry As DirectoryEntry) As Boolean
        Dim flags As Integer = CInt(entry.Properties.Item("UserFlags").Value)
        If (flags And ADApi.UF_ACCOUNTDISABLE) > 0 Then
            Return True
        End If
    End Function

    Public Shared Function MemberOf(ByVal user As DirectoryEntry, ByVal group As String) As Boolean
        For Each o As Object In CType(user.Invoke("Groups", Nothing), System.Collections.IEnumerable)
            Dim g As New DirectoryEntry(o)
            If g.Name.ToLower() = group.ToLower() Then
                Return True
            End If
        Next
        Return False
    End Function

    Public Shared Function GetSid(ByVal entry As DirectoryEntry) As Byte()
        Try
            Dim sid As Byte() = CType(entry.Properties("objectSID").Value, Byte())
            Return sid
        Catch ex As Exception
            Throw New ApplicationException("An error occurred while binding to the group in Active Directory.", ex)
        End Try
    End Function

End Class
