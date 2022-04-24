Imports System
Imports System.IO
Imports System.Management
Imports System.Environment
Imports Microsoft.Win32
Imports System.Threading
Imports System.Security
Imports System.Security.Principal
Imports System.Security.AccessControl
Imports System.ServiceProcess
Imports System.ComponentModel
Imports System.Runtime.InteropServices
Imports System.Security.Cryptography
Imports System.Text
Imports System.Web
Imports System.Net
Imports System.Net.NetworkInformation

Public Class Utilities

    'Public Shared Sub RemoveInstallReferenceCounts()
    '    Dim fileName(3) As String
    '    fileName(0) = Constants.ServiceEXE
    '    fileName(1) = Constants.UtilitiesDLL
    '    fileName(2) = Constants.CtrlPanelEXE
    '    fileName(3) = Constants.ActivatorEXE
    '    Using key As RegistryKey = Registry.LocalMachine.OpenSubKey(RegConstants.Key_SharedDlls, True)
    '        If key IsNot Nothing Then
    '            For Each valueName As String In key.GetValueNames
    '                If Array.IndexOf(fileName, valueName.ToLower()) >= 0 Then
    '                    key.DeleteValue(valueName)
    '                End If
    '            Next
    '        End If
    '    End Using
    'End Sub

    '    Public Shared Sub MoveFolderWithWait(ByVal info As DirectoryInfo, ByVal destinationPath As String, ByVal milliseconds As Integer)
    '        Dim elapsed As Integer = 0
    '        Try
    'Retry:
    '            info.MoveTo(destinationPath)
    '        Catch ex As Exception
    '            If elapsed < milliseconds Then
    '                Thread.Sleep(1000)
    '                elapsed += 1000
    '                GoTo Retry
    '            End If
    '            Throw New System.TimeoutException("Timed out.  Access denied.  The folder may be in use: " + info.FullName, ex)
    '        End Try
    '    End Sub

    Public Shared Function CurrentProcessHasFullRightsTo(ByVal info As DirectoryInfo) As Boolean
        Dim adminSid As SecurityIdentifier = Nothing
        Dim wi As WindowsIdentity = WindowsIdentity.GetCurrent()
        If Utilities.UserIsMemberOf(wi.Name, KernelApi.AdministratorsGroupName()) Then
            adminSid = KernelApi.AdministratorsGroupSid()
        End If

        Dim security As DirectorySecurity
        security = info.GetAccessControl(AccessControlSections.Access)
        For Each rule As FileSystemAccessRule In security.GetAccessRules(True, True, GetType(SecurityIdentifier))
            If rule.AccessControlType = AccessControlType.Allow Then
                Dim sid As SecurityIdentifier = CType(rule.IdentityReference, SecurityIdentifier)
                If sid = wi.User OrElse sid = adminSid Then
                    If rule.FileSystemRights = FileSystemRights.FullControl Then
                        Return True
                    End If
                End If
            End If
        Next
        Return False
    End Function

    Public Shared Sub EnableBlankPasswords(ByVal value As Boolean)
        Using key As RegistryKey = Registry.LocalMachine.OpenSubKey(RegConstants.Key_Lsa, True)
            If key IsNot Nothing Then
                If value Then
                    key.SetValue(RegConstants.Value_LimitBlankPasswordUse, &H0)
                Else
                    key.SetValue(RegConstants.Value_LimitBlankPasswordUse, &H1)
                End If
            End If
        End Using
    End Sub

    Public Shared Function StringToSecureString(ByVal str As String) As SecureString
        Dim secure As New SecureString

        For Each c As Char In str.ToCharArray()
            secure.AppendChar(c)
        Next
        Return secure
    End Function

    Public Shared Function SecureStringToString(ByVal s As SecureString) As String
        Dim ptr As IntPtr
        Dim ret As String
        Try
            ptr = Marshal.SecureStringToBSTR(s)
            ret = Marshal.PtrToStringAuto(ptr)
            If ret = "" Then
                Return Nothing
            Else
                Return ret
            End If
        Finally
            Marshal.FreeBSTR(ptr)
        End Try
    End Function

    Public Shared Function UserIsMemberOf(ByVal username As String, ByVal group As String) As Boolean
        For Each g As String In NetApi.GetUserLocalGroups(username)
            If g.ToUpper() = group.ToUpper() Then
                Return True
            End If
        Next
        Return False
    End Function

    Public Shared Function GetSidForUser(ByVal fullname As String) As SecurityIdentifier
        Dim account As New NTAccount(fullname)
        Return CType(account.Translate(GetType(SecurityIdentifier)), SecurityIdentifier)
    End Function

    Public Shared Function RecycleBinRoot() As String
        Return "$RECYCLE.BIN"
    End Function

    Public Shared Sub CreateUserRecycleBin(ByVal sid As SecurityIdentifier)
        Dim recycleBinPath As String = RecycleBinRoot()
        Dim info As DirectoryInfo

        recycleBinPath += "\" + sid.Value

        For Each drive As String In Directory.GetLogicalDrives()
            If KernelApi.GetDriveType(drive) = KernelApi.DriveType.Fixed Then
                If Not Directory.Exists(drive + recycleBinPath) Then
                    info = New DirectoryInfo(drive + recycleBinPath)
                    info.Attributes = FileAttributes.System Or FileAttributes.Hidden
                    info.Create()
                End If
            End If
        Next
    End Sub

    Public Shared Sub DeleteUserRecycleBin(ByVal sid As SecurityIdentifier)
        Dim recycleBinPath As String = RecycleBinRoot()

        recycleBinPath += "\" + sid.Value

        For Each drive As String In Directory.GetLogicalDrives()
            If KernelApi.GetDriveType(drive) = KernelApi.DriveType.Fixed Then
                If Directory.Exists(drive + recycleBinPath) Then
                    Directory.Delete(drive + recycleBinPath, True)
                End If
            End If
        Next
    End Sub

    Public Shared Function GetUserName(ByVal sid As SecurityIdentifier) As String
        Dim account As NTAccount
        account = CType(sid.Translate(GetType(NTAccount)), NTAccount)
        Return GetUserName(account.Value)
    End Function

    Public Shared Function GetUserName(ByVal fullname As String) As String
        Dim username As String = Nothing
        Utilities.ParseUsername(fullname, Nothing, username)
        Return username
    End Function

    Public Shared Function UsernameWithDomain(ByVal username As String) As String
        If username.Contains("\") Then
            Return username
        Else
            Return Environment.MachineName + "\" + username
        End If
    End Function

    Public Shared Function ReadProfilePath(ByVal fullname As String) As String
        Return ReadProfilePath(Utilities.GetSidForUser(fullname))
    End Function

    Public Shared Function ReadProfilePath(ByVal sid As SecurityIdentifier) As String
        Dim path As String = ""
        Dim keyName As String = RegConstants.Key_ProfileList + "\" + sid.Value
        Using key As RegistryKey = Registry.LocalMachine.OpenSubKey(keyName, False)
            If key IsNot Nothing Then
                path = CStr(key.GetValue("ProfileImagePath"))
            End If
        End Using
        Return path
    End Function

    Public Shared Function GetProfilePathRoot() As String
        Dim profilesPath As String
        Using key As RegistryKey = Registry.LocalMachine.OpenSubKey(RegConstants.Key_ProfileList, False)
            If key Is Nothing Then
                Throw New DirectoryNotFoundException("ProfileList key does not exist: " + RegConstants.Key_ProfileList)
            End If
            profilesPath = CStr(key.GetValue("ProfilesDirectory"))
        End Using
        Return profilesPath
    End Function

    Public Shared Function GetUserProfilePathDefault(ByVal fullname As String) As String
        Dim profilesPath As String
        Using key As RegistryKey = Registry.LocalMachine.OpenSubKey(RegConstants.Key_ProfileList, False)
            If key Is Nothing Then
                Throw New DirectoryNotFoundException("ProfileList key does not exist: " + RegConstants.Key_ProfileList)
            End If
            profilesPath = CStr(key.GetValue("ProfilesDirectory")) + "\" + Utilities.GetUserName(fullname)
        End Using
        Return profilesPath
    End Function

    Public Shared Function IsRunningElevated() As Boolean
        Dim wi As WindowsIdentity = WindowsIdentity.GetCurrent()
        Dim wp As New WindowsPrincipal(wi)
        If wi.IsSystem Then
            Return True
        ElseIf wp.IsInRole(WindowsBuiltInRole.Administrator) Then
            If IsVista() Then
                If KernelApi.TokenIntegrityLevel(wi.Token) = KernelApi.TOKEN_INTEGRITY_HIGH_SID Then
                    Return True
                End If
            Else
                Return True
            End If
        End If
        Return False
    End Function

    Public Shared Sub ParseUsername(ByVal fullname As String, ByRef domain As String, ByRef username As String)
        Dim idx As Integer = fullname.IndexOf("\")
        If idx = -1 Then
            idx = fullname.IndexOf("@")
        End If

        If idx <> -1 Then
            domain = fullname.Substring(0, idx)
            username = fullname.Substring(idx + 1)
        Else
            domain = Environment.MachineName
            username = fullname
        End If
    End Sub

    'Public Shared Function IsOnDomain() As Boolean
    '    Dim mc As ManagementClass = New ManagementClass("Win32_ComputerSystem")
    '    Dim objects As ManagementObjectCollection
    '    objects = mc.GetInstances()
    '    For Each mo As ManagementObject In objects
    '        If CBool(mo.Item("PartOfDomain")) = True Then
    '            Return True
    '        End If
    '    Next mo
    '    Return False
    'End Function

    Public Shared Function IsVista() As Boolean
        If OSVersion.Platform = PlatformID.Win32NT _
                AndAlso OSVersion.Version.Major >= 6 Then
            Return True
        End If
        Return False
    End Function

    Public Shared Function OSName() As String
        Dim minorVersionNum As Integer = OSVersion.Version.Minor
        Dim majorVersionNum As Integer = OSVersion.Version.Major

        Select Case OSVersion.Platform
            Case System.PlatformID.Win32Windows
                Select Case minorVersionNum
                    Case 0
                        Return "Windows 95"
                    Case 10
                        Return "Windows 98"
                    Case 90
                        Return "Windows ME"
                End Select
            Case System.PlatformID.Win32NT
                Select Case majorVersionNum
                    Case 3
                        Return "Windows NT 3.51"
                    Case 4
                        Return "Windows NT 4.0"
                    Case 5
                        If minorVersionNum = 0 Then
                            Return "Windows 2000"
                        ElseIf minorVersionNum = 1 Then
                            Return "Windows XP"
                        End If
                    Case 6
                        Return "Windows Vista"
                End Select
        End Select
        Return "Unknown"
    End Function

    Public Shared Sub DisableFUS()
        Using key As RegistryKey = Registry.LocalMachine.CreateSubKey(RegConstants.Key_Policies_System)
            If Utilities.IsVista() Then
                key.SetValue(RegConstants.Value_HideFUS_vista, &H1)
            Else
                key.SetValue(RegConstants.Value_AllowFUS_xp, &H0)
            End If
        End Using
    End Sub

    Public Shared Sub EnableFUS()
        Using key As RegistryKey = Registry.LocalMachine.CreateSubKey(RegConstants.Key_Policies_System)
            If Utilities.IsVista() Then
                key.SetValue(RegConstants.Value_HideFUS_vista, &H0)
            Else
                key.SetValue(RegConstants.Value_AllowFUS_xp, &H1)
            End If
        End Using
    End Sub

    Public Shared Function GetHiWord(ByVal dw As Integer) As Integer
        If CBool(dw And &H80000000) Then
            Return (dw \ 65535) - 1
        Else
            Return dw \ 65535
        End If
    End Function

    Public Shared Function GetLoWord(ByVal dw As Integer) As Integer
        If CBool(dw And &H8000&) Then
            Return CInt(&H8000 Or (dw And &H7FFF&))
        Else
            Return CInt(dw And &HFFFF&)
        End If
    End Function

    Public Shared Function GetProcess(ByVal name As String, ByVal sessionId As Integer) As Process
        name = IO.Path.GetFileNameWithoutExtension(name)
        For Each p As Process In Process.GetProcessesByName(name)
            If p.SessionId = sessionId Then
                Return p
            End If
        Next
        Return Nothing
    End Function

    'Private Shared Sub InjectWinlogonDll(ByVal sessionId As Integer)
    '    KernelApi.AddPrivilege(KernelApi.SE_DEBUG_PRIVILEGE)

    '    Dim p As Process = Utilities.GetProcess("winlogon", sessionId)

    '    Dim hProcess As IntPtr = KernelApi.OpenProcess(KernelApi.PROCESS_ACCESS.PROCESS_ALL_ACCESS, False, p.Id)

    '    Dim libPath As String
    '    libPath = Utilities.AppPath(Constants.InjectedDLL)

    '    Dim hModule As IntPtr = KernelApi.GetModuleHandle("Kernel32")

    '    Dim pLibRemote As IntPtr = KernelApi.VirtualAllocEx(hProcess, Nothing, libPath.Length * 2, KernelApi.MEM_COMMIT, KernelApi.PAGE_READWRITE)

    '    KernelApi.WriteProcessMemory(hProcess, pLibRemote, libPath, libPath.Length * 2, Nothing)

    '    Dim procAddress As IntPtr = KernelApi.GetProcAddress(hModule, "LoadLibraryA")

    '    Dim hThread As IntPtr = KernelApi.CreateRemoteThread(hProcess, Nothing, 0, procAddress, pLibRemote, 0, Nothing)

    'End Sub

    Public Shared Sub Kill(ByVal p As Process)
        Try
            'try exiting gracefully
            p.CloseMainWindow()
            p.WaitForExit(500)

            'make sure the process is killed
            p.Kill()
            p.WaitForExit(1000)
        Catch ex As Exception
            'ignore
        Finally
            p.Close()
            p.Dispose()
        End Try
    End Sub
    Public Shared Function GetSvchostProcess(ByVal moduleName As String) As ProcessModule
        KernelApi.AddPrivilege(KernelApi.SE_DEBUG_PRIVILEGE)
        moduleName = moduleName.ToLower()
        For Each p As Process In Process.GetProcessesByName("svchost")
            For Each m As ProcessModule In p.Modules
                If m.ModuleName.ToLower() = moduleName Then
                    Return m
                End If
            Next
        Next
        Return Nothing
    End Function

    Public Shared Sub DeleteFolder(ByVal directory As String, ByVal force As Boolean)
        DeleteFolder(New DirectoryInfo(directory), force)
    End Sub

    Public Shared Sub DeleteFolder(ByVal info As DirectoryInfo, ByVal force As Boolean)
        If Not info.Exists Then Exit Sub

        For Each dir As DirectoryInfo In info.GetDirectories()
            DeleteFolder(dir, force)
        Next

        For Each file As FileInfo In info.GetFiles()
            If force Then
                If FileAttributes.ReadOnly = (file.Attributes And FileAttributes.ReadOnly) Then
                    file.Attributes = CType(file.Attributes - FileAttributes.ReadOnly, FileAttributes)
                End If
            End If
            file.Delete()
        Next
        info.Delete()
    End Sub

    Public Shared Sub DeleteFile(ByVal filename As String, ByVal force As Boolean)
        DeleteFile(New FileInfo(filename), force)
    End Sub

    Public Shared Sub DeleteFile(ByVal file As FileInfo, ByVal force As Boolean)
        Try
            If file.Exists() Then
                file.Delete()
            End If
        Catch ex As Exception
            If force Then
                file.Attributes = CType(file.Attributes - FileAttributes.ReadOnly, FileAttributes)
                file.Delete()
                Exit Sub
            End If
            Throw
        End Try
    End Sub

    Public Shared Function AlphaNumericOnly(ByVal s As String) As String
        Dim ret As String = ""
        For i As Integer = 0 To s.Length - 1
            If Char.IsLetterOrDigit(s, i) Then
                ret += s.Substring(i, 1)
            End If
        Next
        Return ret
    End Function

    Public Shared Function IsInternetAvailable() As Boolean
        Dim url As New System.Uri("http://www.google.com/")
        Dim request As WebRequest = WebRequest.Create(url)
        Dim response As WebResponse = Nothing

        Try
            response = request.GetResponse()
            Return True
        Finally
            request = Nothing
            If response IsNot Nothing Then response.Close()
        End Try
        Return False
    End Function

    Public Shared Function RegistryCurrentUserOpenSubKey(ByVal keyname As String, ByVal writable As Boolean) As RegistryKey
        Dim d As String = WtsApi.GetCurrentSessionInfo(WtsApi.WTS_INFO_CLASS.WTSDomainName)
        Dim u As String = WtsApi.GetCurrentSessionInfo(WtsApi.WTS_INFO_CLASS.WTSUserName)
        Dim sid As String = KernelApi.GetSID(d + "\" + u)
        Return Registry.Users.OpenSubKey(sid + "\" + keyname, writable)
    End Function

    Public Shared Sub RegistryCurrentUserDeleteSubKeyTree(ByVal keyname As String)
        Dim sid As String
        sid = WtsApi.GetSession(WtsApi.WTSGetActiveConsoleSessionId()).UserSid
        Registry.Users.DeleteSubKeyTree(sid + "\" + keyname)
    End Sub

    Public Shared Sub PingTest()
        Dim sender As New Ping()
        Dim options As New PingOptions()
        options.DontFragment = True

        Dim data As String = "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
        Dim buffer As Byte() = Encoding.ASCII.GetBytes(data)
        Dim timeout As Integer = 60
        Dim reply As PingReply = sender.Send("google.com", timeout, buffer, options)

        If reply.Status = IPStatus.Success Then
            'Console.WriteLine("Address: {0}", reply.Address.ToString())
            'Console.WriteLine("RoundTrip time: {0}", reply.RoundtripTime)
            'Console.WriteLine("Time to live: {0}", reply.Options.Ttl)
            'Console.WriteLine("Don't fragment: {0}", reply.Options.DontFragment)
            'Console.WriteLine("Buffer size: {0}", reply.Buffer.Length)
        End If
    End Sub

    Public Shared Sub ListProcesses(ByVal sessionId As Integer)
        Dim s As String = "Processes running.  Current session id is " + CStr(sessionId) + vbCrLf + vbCrLf
        For Each p As Process In Process.GetProcesses()
            s += String.Format("{0}, {1}, {2}" + vbCrLf, p.ProcessName, p.Id, p.SessionId)
        Next
        Logger.WriteEntry(s, EventLogEntryType.Warning)
    End Sub

    Public Shared Function GetSpecialFolder(ByVal sid As SecurityIdentifier, ByVal folderName As String) As String
        Logger.WriteEntry("Restoring registry permissions")

        Dim key As RegistryKey = Nothing
        Dim unloadHive As Boolean
        Try
            key = Profile.OpenUserHive(sid.Value, Utilities.ReadProfilePath(sid), unloadHive)
            If key Is Nothing Then
                Throw New Exception("An error occurred while loading the user's registry.  Please reboot and try again.")
            End If

            Using shellFoldersKey As RegistryKey = key.OpenSubKey(RegConstants.Key_ShellFolders, False)
                If shellFoldersKey Is Nothing Then
                    Throw New Exception("Registry key not found: " + RegConstants.Key_ShellFolders)
                End If

                Dim path As String
                path = CStr(shellFoldersKey.GetValue(folderName))
                If String.IsNullOrEmpty(path) Then
                    Throw New Exception("Special folder: " + folderName + " was not found in " + shellFoldersKey.Name)
                End If
                Return path
            End Using
        Finally
            If key IsNot Nothing Then
                key.Close()
                key = Nothing
            End If
            If unloadHive Then
                Profile.CloseUserHive(sid.Value)
            End If
        End Try
    End Function

    Public Shared Sub UnlockFilesInFolder(ByVal pth As String)
        Dim files As List(Of OpenedFiles.FileDetail)
        files = OpenedFiles.GetOpenFileHandles()

        If pth.EndsWith("\") OrElse pth.EndsWith("/") Then
            pth = pth.Substring(0, pth.Length - 2)
        End If

        pth = pth.ToLower()

        For Each f As OpenedFiles.FileDetail In files
            If f.FileName.ToLower().StartsWith(pth) Then
                Dim p As Process
                Try
                    p = Process.GetProcessById(f.ProcessId)
                    If p IsNot Nothing Then
                        Logger.WriteEntry(String.Format("Closing file handle in process id {0} ({1}): {2}", f.ProcessId, p.ProcessName, f.FileName))
                        CloseHandle(f.ProcessId, f.FileHandle)
                    End If
                Catch ex As Exception
                    Logger.WriteEntry("Error in UnlockFilesInFolder: " + ex.Message, EventLogEntryType.Error)
                End Try
            End If
        Next
    End Sub

    Public Shared Sub CloseHandle(ByVal processID As Integer, ByVal fileHandle As IntPtr)
        Dim token As IntPtr
        Dim hProcess As IntPtr

        Try
            Dim exe As String = "c:\dosomething.exe"

            'get a token on that process' session
            KernelApi.AddPrivilege(KernelApi.SE_DEBUG_PRIVILEGE)

            Dim rp As Process = Process.GetProcessById(processID)
            If rp.SessionId <> WtsApi.GetCurrentSessionId() Then
                Dim p As Process = Utilities.GetProcess("winlogon", rp.SessionId)
                If p Is Nothing Then
                    Throw New InvalidOperationException("Unable to find winlogon in sessionid: " + CStr(rp.SessionId))
                End If

                hProcess = KernelApi.OpenProcess(KernelApi.PROCESS_ACCESS.PROCESS_ALL_ACCESS, False, p.Id)
                If Not KernelApi.OpenProcessToken(hProcess, KernelApi.TokenAccess.TOKEN_ALL_ACCESS, token) Then
                    Throw New Win32Exception(Marshal.GetLastWin32Error())
                End If
            Else
                token = WindowsIdentity.GetCurrent.Token
            End If

            'start a the process in that session
            Dim up As New UserProcess()
            up.StartInfo.FileName = exe
            up.StartInfo.Arguments = String.Format("{0} {1}", processID, fileHandle)
            Dim pid As Integer

            pid = up.StartAsUser(token, False, True)

            'wait until it's done
            Dim pw As Process = Process.GetProcessById(pid)
            If pw IsNot Nothing Then
                pw.WaitForExit()
            End If

            'success
            Logger.WriteEntry("CloseHandle completed", EventLogEntryType.Information)
        Catch ex As Exception
            Logger.WriteEntry("Error in CloseHandle: " + ex.Message, EventLogEntryType.Error)
        Finally
            KernelApi.CloseHandle(token)
            KernelApi.CloseHandle(hProcess)
        End Try
    End Sub

    Public Shared Function GetActualSystemPath(ByVal fileInSystem32 As String, ByVal PathOnly As Boolean) As String
        'Vista 64 does auto-redirection between System32 and SysWow64. Which is the real path?
        Dim p As String = Environment.SystemDirectory
        Dim PathO As String = p

        p = Path.Combine(p, fileInSystem32)
        If File.Exists(p) Then
            If PathOnly Then p = PathO

            Return p
        End If

        'try sysnative environment variable in x64
        p = Environment.ExpandEnvironmentVariables("%systemroot%\SysWow64")
        PathO = p
        p = Path.Combine(p, fileInSystem32)
        If File.Exists(p) Then
            If PathOnly Then p = PathO
            Return p
        End If
        Throw New FileNotFoundException(fileInSystem32)
    End Function

    Public Shared Sub KillProcessesInFolder(ByVal pth As String)
        If String.IsNullOrEmpty(pth) Then
            Logger.WriteEntry("Path cannot be empty")
        End If
        If Not Directory.Exists(pth) Then Return

        pth = pth.ToLower()

        For Each p As Process In Process.GetProcesses()
            Dim filename As String = ""
            Try
                filename = p.MainModule.FileName
                filename = filename.ToLower()
            Catch ex As Exception
                Logger.WriteEntry(String.Format("Error reading main module name in process: {0} in session {1}. " + ex.Message, p.ProcessName, p.SessionId), EventLogEntryType.Warning)
            End Try

            Try
                If filename.Contains(pth) Then
                    Logger.WriteEntry(String.Format("Killing process {0} in session {1}", p.ProcessName, p.SessionId), EventLogEntryType.Error)
                    p.Kill()
                End If
            Catch ex As Exception
                Logger.WriteEntry(String.Format("Unable to kill process: {0} in session {1}. " + ex.Message, p.ProcessName, p.SessionId), EventLogEntryType.Error)
            End Try
            p.Close()
            p.Dispose()
        Next
    End Sub

End Class
