Imports System
Imports System.IO
Imports System.Text
Imports System.Management
Imports System.Environment
Imports Microsoft.Win32
Imports System.Threading
Imports System.Security.Principal
Imports System.Security.AccessControl
Imports System.ComponentModel
Imports System.Runtime.InteropServices
Imports System.Security

Public Class Profile

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)> _
    Public Structure PROFILEINFO
        Public dwSize As Integer
        Public dwFlags As Integer
        Public lpUserName As String
        Public lpProfilePath As String
        Public lpDefaultPath As String
        Public lpServerName As String
        Public lpPolicyPath As String
        Public hProfile As Integer
    End Structure

    <DllImport("userenv.dll", SetLastError:=True, CharSet:=CharSet.Unicode)> _
    Private Shared Function LoadUserProfile( _
        ByVal hUserToken As IntPtr, _
        ByRef lpProfileInfo As PROFILEINFO) As Integer
    End Function

    <DllImport("userenv.dll", SetLastError:=True, CharSet:=CharSet.Unicode)> _
    Private Shared Function UnloadUserProfile( _
        ByVal hUserToken As IntPtr, _
        ByVal hProfile As IntPtr) As Integer
    End Function

    Public Sub New()
    End Sub

    Public Shared Function CreateUserProfile(ByVal username As String, ByVal password As String) As String
        Dim token As IntPtr = IntPtr.Zero
        Dim result As Integer
        Dim profilePath As String

        Try
            token = KernelApi.LogonUser(username, password, KernelApi.LogonType.LOGON32_LOGON_NETWORK)

            'if this key exists, LoadUserProfile renames it to bak.  so delete it first
            DeleteProfileListRegistryKey(Utilities.GetSidForUser(username))

            Dim pi As New PROFILEINFO()
            pi.lpUserName = username
            pi.dwSize = Marshal.SizeOf(GetType(PROFILEINFO))

            result = LoadUserProfile(token, pi)
            If result = 0 Then
                Throw New Win32Exception(Marshal.GetLastWin32Error())
            End If

            result = UnloadUserProfile(token, CType(pi.hProfile, IntPtr))
            If result = 0 Then
                Throw New Win32Exception(Marshal.GetLastWin32Error())
            End If

            profilePath = Utilities.ReadProfilePath(username)
            Return profilePath
        Finally
            If token <> IntPtr.Zero Then
                KernelApi.CloseHandle(token)
            End If
        End Try
    End Function


    Private Shared Sub UpdateCredentialFolders(ByVal profilePath As String, ByVal sid As SecurityIdentifier)
        Dim name As String
        Dim info As DirectoryInfo
        Dim dirs() As DirectoryInfo

        If Utilities.IsVista() Then
            name = "AppData\Roaming\Microsoft\Protect"
            info = New DirectoryInfo(Path.Combine(profilePath, name))
            If info.Exists Then
                dirs = info.GetDirectories("S-1-5*")
                If dirs.Length > 0 Then
                    dirs(dirs.Length - 1).MoveTo(Path.Combine(info.FullName, sid.Value))
                End If
            End If
        Else
            name = "Application Data\Microsoft\Protect"
            info = New DirectoryInfo(Path.Combine(profilePath, name))
            If info.Exists Then
                dirs = info.GetDirectories("S-1-5*")
                If dirs.Length > 0 Then
                    dirs(dirs.Length - 1).MoveTo(Path.Combine(info.FullName, sid.Value))
                End If
            End If

            name = "Application Data\Microsoft\Credentials"
            info = New DirectoryInfo(Path.Combine(profilePath, name))
            If info.Exists Then
                dirs = info.GetDirectories("S-1-5*")
                If dirs.Length > 0 Then
                    dirs(dirs.Length - 1).MoveTo(Path.Combine(info.FullName, sid.Value))
                End If
            End If

            name = "Local Settings\Application Data\Microsoft\Credentials"
            info = New DirectoryInfo(Path.Combine(profilePath, name))
            If info.Exists Then
                dirs = info.GetDirectories("S-1-5*")
                If dirs.Length > 0 Then
                    dirs(dirs.Length - 1).MoveTo(Path.Combine(info.FullName, sid.Value))
                End If
            End If
        End If
    End Sub

    Public Shared Function OpenUserHive(ByVal keyname As String, ByVal profilePath As String, ByRef unloadHive As Boolean) As RegistryKey
        unloadHive = False
        Dim key As RegistryKey = Registry.Users.OpenSubKey(keyname, True)
        If key Is Nothing Then
            'not loaded, so load from file
            Dim regFile As String = profilePath + "\NTUSER.DAT"
            If Not File.Exists(regFile) Then
                Logger.WriteEntry("Error loading user's hive.  File not found: " + profilePath, EventLogEntryType.Error)
                Throw New ArgumentException("User's registry hive not found: " + profilePath)
            End If

            KernelApi.AddPrivilege(KernelApi.SE_BACKUP_NAME)
            KernelApi.AddPrivilege(KernelApi.SE_RESTORE_NAME)
            Dim result As Integer
            result = AclApi.RegLoadKey(AclApi.Hives.HKEY_USERS, keyname, profilePath + "\NTUSER.DAT")
            If result <> 0 Then
                Logger.WriteEntry("Error loading user's hive: " + CStr(Marshal.GetLastWin32Error()), EventLogEntryType.Error)
                Throw New Exception("Unable to load the user's registry hive.  Please reboot and try again.")
            End If
            unloadHive = True
            key = Registry.Users.OpenSubKey(keyname, True)
        End If
        Return key
    End Function

    Public Shared Sub CloseUserHive(ByVal keyname As String)
        Dim result As Integer
        result = AclApi.RegUnLoadKey(AclApi.Hives.HKEY_USERS, keyname)
        If result <> 0 Then
            Logger.WriteEntry("Error unloading user's hive: " + CStr(result), EventLogEntryType.Error)
        End If
    End Sub

    Public Shared Function GetOpenUserHiveByUsername(ByVal username As String) As String
        For Each keyName As String In Registry.Users.GetSubKeyNames()
            Using key As RegistryKey = Registry.Users.OpenSubKey(keyName + "\Volatile Environment")
                If key IsNot Nothing Then
                    Dim value As Object
                    value = key.GetValue("USERNAME")
                    If value IsNot Nothing Then
                        If CStr(value).ToLower() = username.ToLower() Then
                            Return keyName
                        End If
                    End If
                    value = key.GetValue("HOMEPATH")
                    If value IsNot Nothing Then
                        If CStr(value).ToLower().Contains(username.ToLower()) Then
                            Return keyName
                        End If
                    End If
                End If
            End Using
        Next
        Return Nothing
    End Function

    'Public Shared Sub CloseUserHive(ByVal username As String)
    '    'attempt to unload the hive when the username/sid is mismatched
    '    Dim keyName As String = GetOpenUserHiveByUsername(username)
    '    If String.IsNullOrEmpty(keyName) Then
    '        Exit Sub
    '    End If
    '    KernelApi.AddPrivilege(KernelApi.SE_RESTORE_NAME)
    '    KernelApi.AddPrivilege(KernelApi.SE_BACKUP_NAME)
    '    Dim result As Integer
    '    result = AclApi.RegUnLoadKey(AclApi.HKEY_USERS, keyName)
    '    If result = 0 Then
    '        'success
    '    ElseIf result = 1314 Then
    '        Throw New PrivilegeNotHeldException("SE_RESTORE_NAME and SE_BACKUP_NAME are required")
    '    Else
    '        Throw New Win32Exception(Marshal.GetLastWin32Error())
    '    End If
    '    keyName += "_Classes"
    '    For Each name As String In Registry.Users.GetSubKeyNames()
    '        If name = keyName Then
    '            result = AclApi.RegUnLoadKey(AclApi.HKEY_USERS, keyName)
    '            If result <> 0 Then
    '                Throw New AccessViolationException("Unable to load the hive: " + keyName)
    '            End If
    '        End If
    '    Next
    'End Sub

    Private Shared Sub SetProfileRegistryPermissions(ByVal profilePath As String, ByVal sid As SecurityIdentifier)
        Logger.WriteEntry("Restoring registry permissions")

        Dim key As RegistryKey = Nothing
        Dim unloadHive As Boolean
        Try
            key = OpenUserHive(sid.Value, profilePath, unloadHive)
            If key Is Nothing Then
                Throw New Exception("An error occurred while loading the user's registry.  Please reboot and try again.")
            End If

            AclApi.SetFullControl(key, sid, True)
            AclApi.SetFullControl(key, "Software\Policies", sid, True)
            AclApi.SetFullControl(key, "Software\Microsoft\Windows\CurrentVersion\Policies", sid, True)
            AclApi.SetFullControl(key, "Software\Microsoft\Windows\CurrentVersion\Group Policy", sid, True)
            AclApi.SetFullControl(key, "Software\Microsoft\Windows\CurrentVersion\Group Policy\GroupMembership", sid, True)

            If Utilities.IsVista() Then
                AclApi.SetFullControl(key, "Software\Microsoft\Internet Explorer\Toolbar\WebBrowser", sid, True)
                AclApi.SetFullControl(key, "Software\Microsoft\Internet Explorer\LowRegistry\DontShowMeThisDialogAgain", sid, True)
                AclApi.SetFullControl(key, "Software\Microsoft\EventSystem", sid, True)
                AclApi.SetFullControl(key, "Software\Microsoft\EventSystem\{26c409cc-ae86-11d1-b616-00805fc79216}", sid, True)
            End If
        Finally
            If key IsNot Nothing Then
                key.Close()
                key = Nothing
            End If
            If unloadHive Then
                CloseUserHive(sid.Value)
            End If
        End Try
    End Sub

    Public Shared Sub DeleteProfileListRegistryKey(ByVal sid As SecurityIdentifier)
        Using key As RegistryKey = Registry.LocalMachine.OpenSubKey(RegConstants.Key_ProfileList, True)
            If key Is Nothing Then
                Throw New Exception("Registry key not found: " + RegConstants.Key_ProfileList)
            End If
            For Each subKeyName As String In key.GetSubKeyNames()
                If subKeyName Like sid.Value + "*" Then
                    key.DeleteSubKeyTree(subKeyName)
                End If
            Next
        End Using
    End Sub

    Public Shared Sub GroupPolicyDisableForceUnload()
        Using key As RegistryKey = Registry.LocalMachine.OpenSubKey(RegConstants.Key_GP_System, True)
            If key IsNot Nothing Then
                key.DeleteValue("DisableForceUnload", False)
            End If
        End Using
    End Sub

    Private Shared Function GetProfileState(ByVal sid As SecurityIdentifier) As Integer
        Using key As RegistryKey = Registry.LocalMachine.OpenSubKey(RegConstants.Key_ProfileList + "\" + sid.Value, False)
            If key IsNot Nothing Then
                Dim val As Object = key.GetValue("State")
                If val IsNot Nothing AndAlso IsNumeric(val) Then
                    Return CInt(val)
                End If
            End If
            Return 0
        End Using
    End Function

    'Hex Mask   State
    '0001       Profile is mandatory.
    '0002       Update the locally cached profile.
    '0004       New local profile.
    '0008       New central profile.
    '0010       Update the central profile.
    '0020       Delete the cached profile.
    '0040       Upgrade the profile.
    '0080       Using Guest user profile.
    '0100       Using Administrator profile.
    '0200       Default net profile is available and ready.
    '0400       Slow network link identified.
    '0800       Temporary profile loaded.

    Private Shared Sub DeleteRegistryValues(ByVal keyName As String, ByVal searchString As String)
        Using key As RegistryKey = Registry.LocalMachine.OpenSubKey(keyName, True)
            If key IsNot Nothing Then
                For Each name As String In key.GetValueNames()
                    If CStr(key.GetValue(name)).ToUpper().Contains(searchString.ToUpper()) Then
                        key.DeleteValue(name)
                    End If
                Next
            End If
        End Using

        Using key As RegistryKey = Utilities.RegistryCurrentUserOpenSubKey(keyName, True)
            If key IsNot Nothing Then
                For Each name As String In key.GetValueNames()
                    If CStr(key.GetValue(name)).ToUpper().Contains(searchString.ToUpper()) Then
                        key.DeleteValue(name)
                    End If
                Next
            End If
        End Using
    End Sub

    Private Shared Sub DeleteGroupPolicyKeys(ByVal sid As SecurityIdentifier)
        'Group Policy
        Using key As RegistryKey = Registry.LocalMachine.OpenSubKey(RegConstants.Key_GroupPolicy, True)
            If key IsNot Nothing Then
                For Each subKeyName As String In key.GetSubKeyNames()
                    If subKeyName = sid.Value Then
                        key.DeleteSubKeyTree(sid.Value)
                        Exit For
                    End If
                Next
            End If
        End Using

        'Group Policy State
        Using key As RegistryKey = Registry.LocalMachine.OpenSubKey(RegConstants.Key_GroupPolicy + "\State", True)
            If key IsNot Nothing Then
                For Each subKeyName As String In key.GetSubKeyNames()
                    If subKeyName = sid.Value Then
                        key.DeleteSubKeyTree(sid.Value)
                        Exit For
                    End If
                Next
            End If
        End Using

        'Group Policy Status
        Using key As RegistryKey = Registry.LocalMachine.OpenSubKey(RegConstants.Key_GroupPolicy + "\Status", True)
            If key IsNot Nothing Then
                For Each subKeyName As String In key.GetSubKeyNames()
                    If subKeyName = sid.Value Then
                        key.DeleteSubKeyTree(sid.Value)
                        Exit For
                    End If
                Next
            End If
        End Using
    End Sub

    Private Shared Sub ClearLastLogonValues(ByVal username As String)
        Using key As RegistryKey = Registry.LocalMachine.OpenSubKey(RegConstants.Key_Winlogon, True)
            If key Is Nothing Then Exit Sub

            Dim value As String
            value = CStr(key.GetValue("DefaultUserName"))
            If Not String.IsNullOrEmpty(value) AndAlso value.Contains(username) Then
                key.SetValue("DefaultUserName", "")
            End If
            value = CStr(key.GetValue("AltDefaultUserName"))
            If Not String.IsNullOrEmpty(value) AndAlso value.Contains(username) Then
                key.SetValue("AltDefaultUserName", "")
            End If
        End Using
    End Sub

    Private Shared Sub DeleteUserFiles(ByVal username As String)
        Dim info As FileInfo
        Dim path As String
        path = Environment.GetFolderPath(SpecialFolder.CommonApplicationData)
        path += "\Microsoft\User Account Pictures"

        info = New FileInfo(path + "\" + username + ".dat")
        If info.Exists() Then info.Delete()

        info = New FileInfo(path + "\" + username + ".bmp")
        If info.Exists() Then info.Delete()

    End Sub

    Public Shared Sub ConfigureProfile(ByVal profilePath As String, ByVal username As String)
        Dim key As RegistryKey = Nothing
        Dim unloadHive As Boolean
        Dim keyName As String
        Try
            key = OpenUserHive(username, profilePath, unloadHive)
            If key Is Nothing Then
                Throw New Exception("An error occurred while loading the user's registry.  Please reboot and try again.")
            End If

            'save the username with the profile
            Using subKey As RegistryKey = key.CreateSubKey("Volatile Environment")
                subKey.SetValue("USERNAME", username)
            End Using

            'disable the phishing dialog
            keyName = "Software\Microsoft\Internet Explorer\PhishingFilter"
            Using subKey As RegistryKey = key.CreateSubKey(keyName)
                subKey.SetValue("Enabled", 0, RegistryValueKind.DWord)
                subKey.SetValue("ShownVerifyBalloon", 3, RegistryValueKind.DWord)
            End Using

            'disable welcome center
            keyName = "Software\Microsoft\Windows\CurrentVersion\Run"
            Using subKey As RegistryKey = key.OpenSubKey(keyName, True)
                If subKey IsNot Nothing Then
                    subKey.DeleteValue("WindowsWelcomeCenter", False)
                    subKey.DeleteValue("Sidebar", False)
                End If
            End Using

            'disable IE information bar notification
            keyName = "Software\Microsoft\Internet Explorer\InformationBar"
            Using subKey As RegistryKey = key.OpenSubKey(keyName, True)
                If subKey IsNot Nothing Then
                    subKey.SetValue("FirstTime", 0, RegistryValueKind.DWord)
                End If
            End Using

            'set wallpaper
            keyName = "Control Panel\Desktop"
            Using subKey As RegistryKey = key.OpenSubKey(keyName, True)
                If subKey IsNot Nothing Then
                    subKey.SetValue("WallpaperStyle", 2, RegistryValueKind.ExpandString)
                    subKey.SetValue("WallpaperOriginX", 0, RegistryValueKind.DWord)
                    subKey.SetValue("WallpaperOriginY", 0, RegistryValueKind.DWord)
                End If
            End Using

            KernelApi.AddPrivilege(KernelApi.SE_SECURITY_PRIVILEGE)

            Dim p As New Process
            p.StartInfo.FileName = Utilities.GetActualSystemPath("Subinacl.exe", False)
            p.StartInfo.UseShellExecute = False
            p.StartInfo.CreateNoWindow = True
            p.StartInfo.Arguments = _
                "/verbose /subkeyreg ""HKEY_LOCAL_MACHINE\Software\Microsoft\Security Center\Svc"" /grant=BUILTIN\Administrators=F"
            p.Start()
            p.WaitForExit()
            If p.ExitCode = 0 Then
                keyName = "Software\Microsoft\Security Center\Svc\"
                keyName += Utilities.GetSidForUser(username).Value
                Using subKey As RegistryKey = Registry.LocalMachine.CreateSubKey(keyName)
                    subKey.SetValue("EnableNotifications", 0, RegistryValueKind.DWord)
                    subKey.SetValue("EnableNotificationsRef", 1, RegistryValueKind.DWord)
                End Using
            End If
        Finally
            If key IsNot Nothing Then
                key.Close()
                key = Nothing
            End If
            If unloadHive Then
                CloseUserHive(username)
            End If
        End Try
    End Sub

    Public Shared Function GetDesktopFolder(ByVal sid As SecurityIdentifier) As String

        Dim desktopPath As String

        'if looking up the current user, use .NET
        If sid.Value = System.Security.Principal.WindowsIdentity.GetCurrent().Owner.Value Then
            desktopPath = Environment.GetFolderPath(SpecialFolder.DesktopDirectory)
            If desktopPath <> "" AndAlso Directory.Exists(desktopPath) Then
                Return desktopPath
            End If
        End If

        'build the path manually
        Dim profilePath As String = Utilities.ReadProfilePath(sid)
        desktopPath = Path.Combine(profilePath, "Desktop")
        If Directory.Exists(desktopPath) Then
            Return desktopPath
        End If

        'open his registry and read it
        desktopPath = Utilities.GetSpecialFolder(sid, "Desktop")
        If Directory.Exists(desktopPath) Then
            Return desktopPath
        End If
        Throw New DirectoryNotFoundException("Unable to find Desktop folder for: " + Utilities.GetUserName(sid))
    End Function

    Public Shared Function GetFavoritesFolder(ByVal sid As SecurityIdentifier) As String

        Dim favoritesPath As String

        'if looking up the current user, use .NET
        If sid.Value = System.Security.Principal.WindowsIdentity.GetCurrent().Owner.Value Then
            favoritesPath = Environment.GetFolderPath(SpecialFolder.Favorites)
            If favoritesPath <> "" AndAlso Directory.Exists(favoritesPath) Then
                Return favoritesPath
            End If
        End If

        'build the path manually
        Dim profilePath As String = Utilities.ReadProfilePath(sid)
        favoritesPath = Path.Combine(profilePath, "Favorites")
        If Directory.Exists(favoritesPath) Then
            Return favoritesPath
        End If

        'open his registry and read it
        favoritesPath = Utilities.GetSpecialFolder(sid, "Favorites")
        If Directory.Exists(favoritesPath) Then
            Return favoritesPath
        End If
        Throw New DirectoryNotFoundException("Unable to find Favorites folder for: " + Utilities.GetUserName(sid))
    End Function

    Public Shared Function GetStartMenuFolder(ByVal sid As SecurityIdentifier) As String

        Dim startMenuPath As String

        'if looking up the current user, use .NET
        If sid.Value = System.Security.Principal.WindowsIdentity.GetCurrent().Owner.Value Then
            startMenuPath = Environment.GetFolderPath(SpecialFolder.StartMenu)
            If startMenuPath <> "" AndAlso Directory.Exists(startMenuPath) Then
                Return startMenuPath
            End If
        End If

        'build the path manually
        Dim profilePath As String = Utilities.ReadProfilePath(sid)
        If Utilities.IsVista() Then
            startMenuPath = Path.Combine(profilePath, "AppData\Roaming\Microsoft\Windows\Start Menu")
        Else
            startMenuPath = Path.Combine(profilePath, "Start Menu")
        End If
        If Directory.Exists(startMenuPath) Then
            Return startMenuPath
        End If

        'open his registry and read it
        startMenuPath = Utilities.GetSpecialFolder(sid, "Start Menu")
        If Directory.Exists(startMenuPath) Then
            Return startMenuPath
        End If
        Throw New DirectoryNotFoundException("Unable to find Desktop folder for: " + Utilities.GetUserName(sid))
    End Function
End Class
