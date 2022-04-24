Imports System.Runtime.InteropServices
Imports System.ComponentModel
Imports System.Security.Principal

Public Class NetApi

    Const READ_CONTROL As Integer = &H20000
    Const TOKEN_QUERY As Integer = &H8
    Const NERR_TRUE As Integer = 1
    Const RPC_S_SERVER_UNAVAILABLE As Integer = 1722
    Const ERROR_INVALID_LEVEL As Integer = 124

    Const ERROR_ACCESS_DENIED As Integer = 5
    Const ERROR_MORE_DATA As Integer = 234

    Const ERROR_INVALID_PASSWORD As Integer = 86

    Const NERR_Success As Integer = 0
    Const NERR_BASE As Integer = 2100
    Const NERR_UserNotFound As Integer = NERR_BASE + 121
    Const NERR_GroupExists As Integer = NERR_BASE + 123
    Const NERR_UserExists As Integer = NERR_BASE + 124
    Const NERR_NotPrimary As Integer = NERR_BASE + 126
    Const NERR_PasswordTooShort As Integer = NERR_BASE + 145
    Const NERR_InvalidComputer As Integer = NERR_BASE + 251
    Const NERR_InvalidMember As Integer = 1388
    Const ERROR_NO_SUCH_GROUP As Integer = 1319
    Const ERROR_NO_SUCH_MEMBER As Integer = 1387
    Const ERROR_MEMBER_IN_ALIAS As Integer = 1378
    Const ERROR_INVALID_SID As Integer = 1337
    Const NERR_BadUsername As Integer = 2202

    Const ERR_ACCOUNT_CHANGED As Integer = &H0
    Const ERR_ACCOUNT_ALREADY_DISABLED As Integer = -&H10
    Const ERR_ACCOUNT_ALREADY_ENABLED As Integer = -&H20
    Const ERR_ACCOUNT_NOT_FOUND As Integer = -&H40
    Const ERR_UPDATE_NOT_SUCCESSFUL As Integer = -&H80

    Const MAX_PREFERRED_LENGTH As Integer = -1
    Const USER_MAXSTORAGE_UNLIMITED As Integer = -1
    Private Const TIMEQ_FOREVER As Integer = -1
    Private Const DOMAIN_GROUP_RID_USERS As Integer = &H201
    Private Const USER_INFO_LEVEL_0 As Integer = 0
    Private Const USER_INFO_LEVEL_1 As Integer = 1
    Private Const USER_INFO_LEVEL_2 As Integer = 2
    Private Const USER_INFO_LEVEL_3 As Integer = 3

    Private Const USER_PRIV_GUEST As Integer = 0
    Private Const USER_PRIV_USER As Integer = 1
    Private Const USER_PRIV_ADMIN As Integer = 2

    Public Const UF_SCRIPT As Integer = &H1
    Public Const UF_ACCOUNTDISABLE As Integer = &H2
    Public Const UF_HOMEDIR_REQUIRED As Integer = &H8
    Public Const UF_LOCKOUT As Integer = &H10
    Public Const UF_PASSWD_NOTREQD As Integer = &H20
    Public Const UF_PASSWD_CANT_CHANGE As Integer = &H40
    Public Const UF_DONT_EXPIRE_PASSWD As Integer = &H10000
    Public Const UF_NORMAL_ACCOUNT As Integer = &H200
    Public Const UF_SERVER_TRUST_ACCOUNT As Integer = &H2000
    Public Const UF_TEMP_DUPLICATE_ACCOUNT As Integer = &H100
    Public Const UF_INTERDOMAIN_TRUST_ACCOUNT As Integer = &H800
    Public Const UF_WORKSTATION_TRUST_ACCOUNT As Integer = &H1000
    Public Const UF_ENCRYPTED_TEXT_PASSWORD_ALLOWED As Integer = &H80
    Public Const UF_MNS_LOGON_ACCOUNT As Integer = &H20000
    Public Const UF_SMARTCARD_REQUIRED As Integer = &H40000
    Public Const UF_TRUSTED_FOR_DELEGATION As Integer = &H80000
    Public Const UF_NOT_DELEGATED As Integer = &H100000
    Public Const UF_USE_DES_KEY_ONLY As Integer = &H200000
    Public Const UF_DONT_REQUIRE_PREAUTH As Integer = &H400000

    Private Const STILL_ACTIVE As Integer = &H103
    Private Const PROCESS_QUERY_INFORMATION As Integer = &H400

    Public Structure User
        Public Name As String
        Public Comment As String
        Public FullName As String
        Public Flags As Integer
        Public Id As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)> _
    Private Structure LOCALGROUP_MEMBERS_INFO_3
        <MarshalAs(UnmanagedType.LPWStr)> Public FullName As String
    End Structure

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)> _
    Public Structure USER_INFO_0
        Public Username As String
    End Structure

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)> _
    Public Structure USER_INFO_1
        Public Username As String
        Public Password As String
        Public Password_Age As Integer
        Public Priv As Integer
        Public Home_Dir As String
        Public Comment As String
        Public Flags As Integer
        Public Script_Path As String
    End Structure

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)> _
    Public Structure USER_INFO_2
        <MarshalAs(UnmanagedType.LPWStr)> Public username As String
        <MarshalAs(UnmanagedType.LPWStr)> Public password As String
        Public password_age As Integer
        Public priv As Integer
        <MarshalAs(UnmanagedType.LPWStr)> Public home_dir As String
        <MarshalAs(UnmanagedType.LPWStr)> Public comment As String
        Public flags As Integer
        <MarshalAs(UnmanagedType.LPWStr)> Public script_path As String
        Public auth_flags As Integer
        <MarshalAs(UnmanagedType.LPWStr)> Public full_name As String
        <MarshalAs(UnmanagedType.LPWStr)> Public usr_comment As String
        <MarshalAs(UnmanagedType.LPWStr)> Public parms As String
        <MarshalAs(UnmanagedType.LPWStr)> Public workstations As String
        Public last_logon As Integer
        Public last_logoff As Integer
        Public acct_expires As Integer
        Public max_storage As Integer
        Public units_per_week As Integer
        Public logon_hours As IntPtr
        Public bad_pw_count As Integer
        Public num_logons As Integer
        <MarshalAs(UnmanagedType.LPWStr)> Public logon_server As String
        Public country_code As Integer
        Public code_page As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)> _
    Public Structure USER_INFO_3
        <MarshalAs(UnmanagedType.LPWStr)> Public username As String
        <MarshalAs(UnmanagedType.LPWStr)> Public password As String
        Public password_age As Integer
        Public priv As Integer
        <MarshalAs(UnmanagedType.LPWStr)> Public home_dir As String
        <MarshalAs(UnmanagedType.LPWStr)> Public comment As String
        Public flags As Integer
        <MarshalAs(UnmanagedType.LPWStr)> Public script_path As String
        Public auth_flags As Integer
        <MarshalAs(UnmanagedType.LPWStr)> Public full_name As String
        <MarshalAs(UnmanagedType.LPWStr)> Public usr_comment As String
        <MarshalAs(UnmanagedType.LPWStr)> Public parms As String
        <MarshalAs(UnmanagedType.LPWStr)> Public workstations As String
        Public last_logon As Integer
        Public last_logoff As Integer
        Public acct_expires As Integer
        Public max_storage As Integer
        Public units_per_week As Integer
        Public logon_hours As IntPtr
        Public bad_pw_count As Integer
        Public num_logons As Integer
        <MarshalAs(UnmanagedType.LPWStr)> Public logon_server As String
        Public country_code As Integer
        Public code_page As Integer
        Public user_id As Integer
        Public primary_group_id As Integer
        <MarshalAs(UnmanagedType.LPWStr)> Public profile As String
        <MarshalAs(UnmanagedType.LPWStr)> Public home_dir_drive As String
        Public password_expired As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)> _
    Private Structure USER_INFO_1008
        Public flags As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)> _
    Private Structure NET_DISPLAY_USER
        Public name As IntPtr
        Public comment As IntPtr
        Public flags As UInteger
        Public full_name As IntPtr
        Public user_id As UInteger
        Public next_index As UInteger
    End Structure

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)> _
    Private Structure NET_DISPLAY_GROUP
        Public grp_name As IntPtr
        Public grp_comment As IntPtr
        Public grp_group_id As Integer
        Public grp_attributes As Integer
        Public grp_next_index As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)> _
    Private Structure NET_DISPLAY_MACHINE
        Public name As IntPtr
        Public comment As IntPtr
        Public flags As Integer
        Public user_id As Integer
        Public next_index As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)> _
    Private Structure LOCALGROUP_USERS_INFO_0
        <MarshalAs(UnmanagedType.LPWStr)> Public name As String
    End Structure

    <DllImport("netapi32.dll", CharSet:=CharSet.Unicode)> _
    Private Shared Function NetUserChangePassword( _
        <MarshalAs(UnmanagedType.LPWStr)> ByVal Domain As String, _
        <MarshalAs(UnmanagedType.LPWStr)> ByVal Username As String, _
        <MarshalAs(UnmanagedType.LPWStr)> ByVal OldPassword As String, _
        <MarshalAs(UnmanagedType.LPWStr)> ByVal NewPassword As String) As Integer
    End Function

    <DllImport("netapi32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)> _
    Private Shared Function NetUserSetInfo( _
        ByVal servername As String, _
        ByVal username As String, _
        ByVal level As Integer, _
        ByRef buf As String, _
        ByRef parm_err As Integer) As Integer
    End Function

    <DllImport("netapi32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)> _
    Private Shared Function NetUserSetInfo( _
        ByVal servername As String, _
        ByVal username As String, _
        ByVal level As Integer, _
        ByVal ptr As IntPtr, _
        ByRef parm_err As Integer) As Integer
    End Function

    <DllImport("netapi32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)> _
    Private Shared Function NetUserSetInfo( _
        ByVal servername As String, _
        ByVal username As String, _
        ByVal level As Integer, _
        ByRef info As USER_INFO_0, _
        ByRef parm_err As Integer) As Integer
    End Function

    <DllImport("netapi32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)> _
    Private Shared Function NetUserSetInfo( _
        ByVal servername As String, _
        ByVal username As String, _
        ByVal level As Integer, _
        ByRef info As USER_INFO_1, _
        ByRef parm_err As Integer) As Integer
    End Function

    <DllImport("netapi32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)> _
    Private Shared Function NetUserSetInfo( _
        ByVal servername As String, _
        ByVal username As String, _
        ByVal level As Integer, _
        ByRef info As USER_INFO_2, _
        ByRef parm_err As Integer) As Integer
    End Function

    <DllImport("netapi32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)> _
    Private Shared Function NetUserGetInfo( _
        ByVal servername As String, _
        ByVal username As String, _
        ByVal level As Integer, _
        ByRef bufptr As IntPtr) As Integer
    End Function

    <DllImport("netapi32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)> _
    Private Shared Function NetUserGetLocalGroups( _
        <MarshalAs(UnmanagedType.LPWStr)> ByVal servername As String, _
        <MarshalAs(UnmanagedType.LPWStr)> ByVal username As String, _
        ByVal level As Integer, _
        ByVal flags As Integer, _
        ByRef bufptr As IntPtr, _
        ByVal prefmaxlen As Integer, _
        ByRef entriesread As Integer, _
        ByRef totalentries As Integer) As Integer
    End Function

    <DllImport("netapi32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)> _
    Private Shared Function NetUserAdd( _
        <MarshalAs(UnmanagedType.LPWStr)> ByVal ServerName As String, _
        ByVal Level As Integer, _
        ByRef Buffer As USER_INFO_2, _
        ByVal parm_err As Integer) As Integer
    End Function

    <DllImport("netapi32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)> _
    Private Shared Function NetUserAdd( _
        <MarshalAs(UnmanagedType.LPWStr)> ByVal ServerName As String, _
        ByVal Level As Integer, _
        ByRef Buffer As USER_INFO_1, _
        ByVal parm_err As Integer) As Integer
    End Function

    <DllImport("netapi32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)> _
    Private Shared Function NetApiBufferAllocate( _
        ByVal ByteCount As Integer, _
        ByVal Ptr As IntPtr) As Integer
    End Function

    <DllImport("netapi32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)> _
    Private Shared Function NetApiBufferFree(ByVal bufptr As IntPtr) As Integer
    End Function

    <DllImport("netapi32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)> _
    Private Shared Function NetQueryDisplayInformation( _
        <MarshalAs(UnmanagedType.LPWStr)> ByVal ServerName As String, _
        ByVal Level As Integer, _
        ByVal Index As Integer, _
        ByVal EntriesRequested As Integer, _
        ByVal PreferredMaximumLength As Integer, _
        ByRef ReturnedEntryCount As Integer, _
        ByRef SortedBuffer As IntPtr) As Integer
    End Function

    <DllImport("netapi32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)> _
    Private Shared Function NetUserEnum( _
        <MarshalAs(UnmanagedType.LPWStr)> ByVal servername As String, _
        ByVal level As Integer, _
        ByVal filter As Integer, _
        ByRef bufptr As IntPtr, _
        ByVal prefmaxlen As Integer, _
        ByRef entriesread As Integer, _
        ByRef totalentries As Integer, _
        ByRef resume_handle As Integer) As Integer
    End Function

    <DllImport("netapi32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)> _
    Private Shared Function NetUserDel( _
        ByVal ServerName As String, _
        ByVal UserName As String) As Integer
    End Function

    <DllImport("userenv.dll", SetLastError:=True, CharSet:=CharSet.Unicode)> _
    Private Shared Function DeleteProfile( _
        ByVal lpSidString As String, _
        ByVal lpProfilePath As String, _
        ByVal lpComputerName As String) As Integer
    End Function

    <DllImport("netapi32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)> _
    Private Shared Function NetLocalGroupAddMembers( _
        ByVal ServerName As String, _
        ByVal GroupName As String, _
        ByVal Level As UInteger, _
        ByRef Buffer As LOCALGROUP_MEMBERS_INFO_3, _
        ByVal TotalEntries As UInteger) As Integer
    End Function

    Public Shared Sub DeleteUser(ByVal fullname As String, ByVal deleteProfile As Boolean)
        Utilities.ParseUsername(fullname, Nothing, fullname)
        Dim result As Integer
        Dim sid As SecurityIdentifier = Utilities.GetSidForUser(fullname)

        If deleteProfile Then
            Dim profilePath As String = Utilities.ReadProfilePath(sid)
            If Not String.IsNullOrEmpty(profilePath) Then
                Dim retries As Integer = 0
                Try
Retry:
                    result = NetApi.DeleteProfile(sid.Value, Nothing, Nothing)
                    If result = 0 Then
                        Throw New Win32Exception(Marshal.GetLastWin32Error())
                    End If
                Catch ex As Exception
                    retries += 1
                    If retries < 3 Then
                        GoTo retry
                    End If
                    Try
                        If System.IO.Directory.Exists(profilePath) Then
                            System.IO.Directory.Delete(profilePath, True)
                        End If
                    Catch ex2 As Exception
                        Logger.WriteEntry("Unable to delete profile: " + ex2.Message, EventLogEntryType.Error)
                    End Try
                End Try
            End If
        End If

        result = NetUserDel(Nothing, fullname)
        If result <> NERR_Success Then
            Throw New Win32Exception(result)
        End If
    End Sub

    Public Shared Function GetUsers() As List(Of User)
        Dim _users As New List(Of User)
        Dim buf As IntPtr = IntPtr.Zero
        Dim nextIndex As Integer
        Dim entriesRead As Integer

        Dim usr As NET_DISPLAY_USER

        Dim result As Integer = ERROR_MORE_DATA
        Do While result = ERROR_MORE_DATA
            Try
                result = NetQueryDisplayInformation(Nothing, 1, nextIndex, 100, MAX_PREFERRED_LENGTH, entriesRead, buf)
                If result = ERROR_ACCESS_DENIED OrElse result = ERROR_INVALID_LEVEL Then
                    Throw New Win32Exception(Marshal.GetLastWin32Error())
                End If

                Dim users(entriesRead) As NET_DISPLAY_USER
                Dim iter As IntPtr = buf

                For i As Integer = 0 To entriesRead - 1
                    users(i) = CType(Marshal.PtrToStructure(iter, GetType(NET_DISPLAY_USER)), NET_DISPLAY_USER)
                    iter = New IntPtr(iter.ToInt32 + Marshal.SizeOf(GetType(NET_DISPLAY_USER)))

                    Dim u As New User
                    u.Name = Marshal.PtrToStringAuto(users(i).name)
                    u.Flags = CInt(users(i).flags)
                    u.Id = CInt(users(i).user_id)
                    'u.FullName = Marshal.PtrToStringAuto(users(i).full_name)
                    _users.Add(u)
                    nextIndex = CInt(usr.next_index)
                Next i
            Finally
                NetApiBufferFree(buf)
            End Try
        Loop
        Return _users
    End Function

    Public Shared Function UserListEnum() As ArrayList
        Dim _users As New ArrayList
        Dim entriesRead As Integer
        Dim totalEntries As Integer
        Dim resumeHandle As Integer
        Dim bufPtr As IntPtr

        NetUserEnum(Nothing, 0, 2, bufPtr, -1, entriesRead, totalEntries, resumeHandle)
        If entriesRead = 0 Then
            Throw New Win32Exception(Marshal.GetLastWin32Error())
        End If

        Dim users(entriesRead) As USER_INFO_0
        Dim iter As IntPtr = bufPtr
        For i As Integer = 0 To entriesRead - 1
            users(i) = CType(Marshal.PtrToStructure(iter, GetType(USER_INFO_0)), USER_INFO_0)
            iter = New IntPtr(iter.ToInt32 + Marshal.SizeOf(GetType(USER_INFO_0)))

            _users.Add(users(i).Username)
        Next i
        NetApiBufferFree(bufPtr)
        Return _users
    End Function

    Public Shared Function CreateUser(ByVal username As String, ByVal password As String, ByVal isAdmin As Boolean, ByVal fullname As String, ByVal comments As String) As SecurityIdentifier
        Dim flags As Integer = UF_NORMAL_ACCOUNT
        'Dim info As New USER_INFO_2()
        Dim info As New USER_INFO_1()

        Utilities.ParseUsername(username, Nothing, username)

        With info
            .username = username    ' Name of user account
            .password = password    ' Password for user account
            .password_age = 25637   ' Ignored by NetUserAdd
            .priv = USER_PRIV_USER  ' Must be USER_PRIV_USER
            .home_dir = Nothing     ' Can be null
            .comment = comments     ' Can be null or null terminated string
            .flags = UF_DONT_EXPIRE_PASSWD Or UF_NORMAL_ACCOUNT Or UF_SCRIPT   ' There are a number of variations
            .script_path = Nothing  ' Path of user's logon script.  Can be null
            '.auth_flags = 0         ' NetUserAdd must be 0
            '.full_name = fullname   ' User's full name.  Can be null
            '.usr_comment = Nothing  ' User comment.  Can be null
            '.parms = Nothing        ' Used by specific applications.  Can be null
            '.workstations = Nothing ' Workstations a user can log onto (null = all stations)
            '.last_logon = 0         ' Ignored by NetUserAdd
            '.last_logoff = 0        ' Not used
            '.acct_expires = TIMEQ_FOREVER                ' Never expires
            '.max_storage = USER_MAXSTORAGE_UNLIMITED     ' Can use any amount of disk space
            '.units_per_week = 0     ' Ignored by NetUserAdd
            '.logon_hours = IntPtr.Zero  ' Null means no restrictions
            '.bad_pw_count = 0       ' Ignored by NetUserAdd
            '.num_logons = 0         ' Ignored by NetUserAdd
            '.logon_server = Nothing ' Null means logon to domain server
            '.country_code = 0       ' Country code for user's language
            '.code_page = 0          ' Code page for user's language
        End With

        Try

            Dim parmErr As Integer
            Dim result As Integer
            result = NetUserAdd(Nothing, USER_INFO_LEVEL_1, info, parmErr)
            If result <> NERR_Success Then
                Throw New Win32Exception(result)
            End If

            Dim groupInfo3 As LOCALGROUP_MEMBERS_INFO_3
            groupInfo3.FullName = Environment.MachineName + "\" + username

            Dim groupName As String
            If isAdmin Then
                groupName = KernelApi.AdministratorsGroupName()
            Else
                groupName = "Users"
            End If
            result = NetLocalGroupAddMembers(Nothing, groupName, 3, groupInfo3, 1)
            If result <> NERR_Success Then
                Throw New Win32Exception(result)
            End If

            Return Utilities.GetSidForUser(username)
        Finally
            'Marshal.FreeHGlobal(ptr)
        End Try
    End Function

    Public Shared Function UserExists(ByVal username As String) As Boolean
        Try
            NetApi.GetUserInfo0(username)
            Return True
        Catch aex As ArgumentOutOfRangeException
            'user not found
        End Try
        Return False
    End Function

    Public Shared Function GetUserLocalGroups(ByVal username As String) As List(Of String)
        Dim entriesRead As Integer
        Dim totalEntries As Integer
        Dim bufPtr As IntPtr = IntPtr.Zero
        Dim result As Integer

        Try
            result = NetApi.NetUserGetLocalGroups(Nothing, username, 0, 0, bufPtr, 1024, entriesRead, totalEntries)
            Select Case result
                Case NERR_Success
                    'ok
                Case ERROR_ACCESS_DENIED
                    Throw New AccessViolationException("Unable to retrieve groups.  Access denied.")
                Case ERROR_MORE_DATA
                    'ok
                Case NERR_InvalidComputer
                    Throw New ArgumentOutOfRangeException("Unable to retrieve groups. Invalid computer name.")
                Case NERR_UserNotFound
                    Throw New ArgumentOutOfRangeException("username, value is: " + username)
            End Select
            If result <> 0 Then
                Throw New Win32Exception(Marshal.GetLastWin32Error())
            End If

            Dim groupList As New List(Of String)

            If entriesRead > 0 Then
                Dim groups(entriesRead - 1) As LOCALGROUP_USERS_INFO_0
                Dim iter As IntPtr = bufPtr

                For i As Integer = 0 To entriesRead - 1
                    groups(i) = CType(Marshal.PtrToStructure(iter, GetType(LOCALGROUP_USERS_INFO_0)), LOCALGROUP_USERS_INFO_0)
                    iter = CType(iter.ToInt32() + Marshal.SizeOf(GetType(LOCALGROUP_USERS_INFO_0)), IntPtr)
                    groupList.Add(groups(i).name)
                Next

            End If
            Return groupList
        Catch argEx As ArgumentOutOfRangeException
            'username isn't valid
            Return New List(Of String)
        Finally
            If bufPtr <> IntPtr.Zero Then
                NetApiBufferFree(bufPtr)
            End If
        End Try
    End Function

    Public Shared Sub ChangeUsername(ByVal existingName As String, ByVal newName As String)
        Dim username As String = Nothing
        Utilities.ParseUsername(existingName, Nothing, existingName)

        If NetApi.UserExists(newName) Then
            Throw New ArgumentException("Username already exists")
        End If

        Dim info As USER_INFO_0
        info = GetUserInfo0(existingName)
        info.Username = newName

        Dim result As Integer
        result = NetUserSetInfo(Nothing, existingName, USER_INFO_LEVEL_0, info, 0)
        Select Case result
            Case NERR_Success
                '
            Case ERROR_ACCESS_DENIED
                Throw New AccessViolationException("Privilege not held: SE_CHANGE_NOTIFY_NAME")
            Case NERR_InvalidComputer
                Throw New ArgumentOutOfRangeException("Invalid computer name")
            Case NERR_NotPrimary
                Throw New InvalidOperationException("Not valid on primary domain")
            Case NERR_UserNotFound
                Throw New ArgumentOutOfRangeException("User not found")
            Case NERR_BadUsername
                Throw New ArgumentOutOfRangeException("Bad user name: " + newName)
            Case Else
                Throw New Exception("Unknown error while renaming: " + CStr(result))
        End Select
    End Sub

    Public Shared Function GetUserInfo0(ByVal fullname As String) As USER_INFO_0
        Dim username As String = Nothing
        Dim domain As String = Nothing
        Utilities.ParseUsername(fullname, domain, username)

        Dim ptr As IntPtr

        Try
            Dim result As Integer
            result = NetUserGetInfo(Nothing, username, USER_INFO_LEVEL_1, ptr)
            Select Case result
                Case NERR_Success
                    '
                Case ERROR_ACCESS_DENIED
                    Throw New AccessViolationException("Privilege not held: SE_CHANGE_NOTIFY_NAME")
                Case NERR_InvalidComputer
                    Throw New ArgumentOutOfRangeException("Invalid computer name")
                Case NERR_NotPrimary
                    Throw New InvalidOperationException("Not primary domain")
                Case NERR_UserNotFound
                    Throw New ArgumentOutOfRangeException("User not found")
                Case NERR_PasswordTooShort
                    Throw New ArgumentOutOfRangeException("New password does not meet policy restrictions")
                Case NERR_BadUsername
                    Throw New ArgumentOutOfRangeException("Bad user name: " + username)
                Case Else
                    Throw New Exception("Unknown error retrieving GetUserInfo0: " + CStr(result))
            End Select
            Return CType(Marshal.PtrToStructure(ptr, GetType(USER_INFO_0)), USER_INFO_0)
        Finally
            NetApiBufferFree(ptr)
        End Try
    End Function

    Public Shared Function GetUserInfo1(ByVal fullname As String) As USER_INFO_1
        Dim username As String = Nothing
        Dim domain As String = Nothing
        Utilities.ParseUsername(fullname, domain, username)

        Dim ptr As IntPtr

        Try
            Dim result As Integer
            result = NetUserGetInfo(Nothing, username, USER_INFO_LEVEL_1, ptr)
            Select Case result
                Case NERR_Success
                    '
                Case ERROR_ACCESS_DENIED
                    Throw New AccessViolationException("Privilege not held: SE_CHANGE_NOTIFY_NAME")
                Case NERR_InvalidComputer
                    Throw New ArgumentOutOfRangeException("Invalid computer name")
                Case NERR_NotPrimary
                    Throw New InvalidOperationException("Not valid on primary domain")
                Case NERR_UserNotFound
                    Throw New ArgumentOutOfRangeException("User not found")
                Case NERR_PasswordTooShort
                    Throw New ArgumentOutOfRangeException("New password does not meet policy restrictions")
                Case NERR_BadUsername
                    Throw New ArgumentOutOfRangeException("Bad user name: " + username)
                Case Else
                    Throw New Exception("Unknown error retrieving GetUserInfo1: " + CStr(result))
            End Select
            Return CType(Marshal.PtrToStructure(ptr, GetType(USER_INFO_1)), USER_INFO_1)
        Finally
            NetApiBufferFree(ptr)
        End Try
    End Function

    Public Shared Function GetUserInfo2(ByVal fullname As String) As USER_INFO_2
        Dim username As String = Nothing
        Dim domain As String = Nothing
        Utilities.ParseUsername(fullname, domain, username)

        Dim ptr As IntPtr

        Try
            Dim result As Integer
            result = NetUserGetInfo(Nothing, username, USER_INFO_LEVEL_2, ptr)
            Select Case result
                Case NERR_Success
                    '
                Case ERROR_ACCESS_DENIED
                    Throw New AccessViolationException("Privilege not held: SE_CHANGE_NOTIFY_NAME")
                Case NERR_InvalidComputer
                    Throw New ArgumentOutOfRangeException("Invalid computer name")
                Case NERR_NotPrimary
                    Throw New InvalidOperationException("Not valid on primary domain")
                Case NERR_UserNotFound
                    Throw New ArgumentOutOfRangeException("User not found")
                Case NERR_PasswordTooShort
                    Throw New ArgumentOutOfRangeException("New password does not meet policy restrictions")
                Case NERR_BadUsername
                    Throw New ArgumentOutOfRangeException("Bad user name: " + username)
                Case Else
                    Throw New Exception("Unknown error retrieving GetUserInfo2: " + CStr(result))
            End Select
            Return CType(Marshal.PtrToStructure(ptr, GetType(USER_INFO_2)), USER_INFO_2)
        Finally
            NetApiBufferFree(ptr)
        End Try
    End Function

    Public Shared Function GetUserInfo3(ByVal fullname As String) As USER_INFO_3
        Dim username As String = Nothing
        Dim domain As String = Nothing
        Utilities.ParseUsername(fullname, domain, username)

        Dim ptr As IntPtr

        Try
            Dim result As Integer
            result = NetUserGetInfo(Nothing, username, USER_INFO_LEVEL_3, ptr)
            Select Case result
                Case NERR_Success
                    '
                Case ERROR_ACCESS_DENIED
                    Throw New AccessViolationException("Privilege not held: SE_CHANGE_NOTIFY_NAME")
                Case NERR_InvalidComputer
                    Throw New ArgumentOutOfRangeException("Invalid computer name")
                Case NERR_NotPrimary
                    Throw New InvalidOperationException("Not valid on primary domain")
                Case NERR_UserNotFound
                    Throw New ArgumentOutOfRangeException("User not found")
                Case NERR_PasswordTooShort
                    Throw New ArgumentOutOfRangeException("New password does not meet policy restrictions")
                Case NERR_BadUsername
                    Throw New ArgumentOutOfRangeException("Bad user name: " + username)
                Case Else
                    Throw New Exception("Unknown error retrieving GetUserInfo3: " + CStr(result))
            End Select
            Return CType(Marshal.PtrToStructure(ptr, GetType(USER_INFO_3)), USER_INFO_3)
        Finally
            NetApiBufferFree(ptr)
        End Try
    End Function

    Public Shared Function IsDisabled(ByVal fullname As String) As Boolean
        Dim info As USER_INFO_1
        info = GetUserInfo1(fullname)

        If (info.Flags And NetApi.UF_ACCOUNTDISABLE) = NetApi.UF_ACCOUNTDISABLE Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Shared Sub EnableAccount(ByVal fullname As String)
        EnableAccount(fullname, True)
    End Sub

    Public Shared Sub DisableAccount(ByVal fullname As String)
        EnableAccount(fullname, False)
    End Sub

    Private Shared Sub EnableAccount(ByVal fullname As String, ByVal value As Boolean)
        Dim username As String = Nothing
        Dim domain As String = Nothing
        Utilities.ParseUsername(fullname, domain, username)

        Dim info As USER_INFO_1
        info = GetUserInfo1(fullname)

        If value Then
            info.Flags = info.Flags And Not NetApi.UF_ACCOUNTDISABLE
        Else
            info.Flags = info.Flags Or NetApi.UF_ACCOUNTDISABLE
        End If

        Dim result As Integer
        result = NetUserSetInfo(Nothing, username, USER_INFO_LEVEL_1, info, 0)
        Select Case result
            Case NERR_Success
                '
            Case ERROR_ACCESS_DENIED
                Throw New AccessViolationException("Privilege not held: SE_CHANGE_NOTIFY_NAME")
            Case NERR_InvalidComputer
                Throw New ArgumentOutOfRangeException("Invalid computer name")
            Case NERR_NotPrimary
                Throw New InvalidOperationException("Not valid on primary domain")
            Case NERR_UserNotFound
                Throw New ArgumentOutOfRangeException("User not found")
            Case NERR_BadUsername
                Throw New ArgumentOutOfRangeException("Bad user name: " + username)
            Case Else
                Throw New Exception("Unknown error enabling account: " + CStr(result))
        End Select
    End Sub

    Public Shared Sub SetPassword(ByVal fullname As String, ByVal password As String)
        Dim username As String = Nothing
        Dim domain As String = Nothing
        Utilities.ParseUsername(fullname, domain, username)

        Dim result As Integer
        result = NetUserSetInfo(domain, username, 1003, password, 0)
        Select Case result
            Case NERR_Success
                '
            Case ERROR_ACCESS_DENIED
                Throw New AccessViolationException("Privilege not held: SE_CHANGE_NOTIFY_NAME")
            Case ERROR_INVALID_PASSWORD
                Throw New ArgumentException("Invalid password")
            Case NERR_InvalidComputer
                Throw New ArgumentOutOfRangeException("Invalid computer name")
            Case NERR_NotPrimary
                Throw New InvalidOperationException("Not valid on primary domain")
            Case NERR_UserNotFound
                Throw New ArgumentOutOfRangeException("User not found")
            Case NERR_PasswordTooShort
                Throw New ArgumentOutOfRangeException("New password does not meet policy restrictions")
            Case NERR_BadUsername
                Throw New ArgumentOutOfRangeException("Bad user name: " + username)
            Case Else
                Throw New Exception("Unknown error setting password: " + CStr(result))
        End Select
    End Sub

    Public Shared Sub ChangePassword(ByVal fullname As String, ByVal oldPassword As String, ByVal newPassword As String)
        Dim username As String = Nothing
        Dim domain As String = Nothing
        Utilities.ParseUsername(fullname, domain, username)

        Dim result As Integer
        result = NetUserChangePassword(domain, username, oldPassword, newPassword)
        Select Case result
            Case NERR_Success
                '
            Case ERROR_ACCESS_DENIED
                Throw New AccessViolationException("Privilege not held: SE_CHANGE_NOTIFY_NAME")
            Case ERROR_INVALID_PASSWORD
                Throw New ArgumentException("The password is incorrect")
            Case NERR_InvalidComputer
                Throw New ArgumentOutOfRangeException("Invalid computer name")
            Case NERR_NotPrimary
                Throw New InvalidOperationException("Not valid on primary domain")
            Case NERR_UserNotFound
                Throw New ArgumentOutOfRangeException("User not found")
            Case NERR_PasswordTooShort
                Throw New ArgumentOutOfRangeException("New password does not meet policy restrictions")
            Case NERR_BadUsername
                Throw New ArgumentOutOfRangeException("Bad user name: " + username)
            Case Else
                Throw New Exception("Unknown error changing password: " + CStr(result))
        End Select
    End Sub

End Class
