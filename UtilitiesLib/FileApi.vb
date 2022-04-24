Imports System.ComponentModel
Imports System.Runtime.InteropServices
Imports Microsoft.Win32
Imports Microsoft.Win32.SafeHandles
Imports System.IO

Public Class FileApi
    Private Const GENERIC_READ As Integer = &H80000000
    Private Const GENERIC_WRITE As Integer = &H40000000
    Private Const READ_CONTROL As Integer = &H20000
    Private Const FILE_SHARE_READ As Integer = &H1
    Private Const FILE_SHARE_DELETE As Integer = &H4
    Private Const FILE_FLAG_BACKUP_SEMANTICS As Integer = &H2000000
    Private Const OPEN_EXISTING As Integer = 3

    Public Enum FileAccessEnum As Integer
        GenericRead = &H80000000
        GenericWrite = &H40000000
        GenericExecute = &H20000000
        GenericAll = &H10000000
    End Enum

    Public Enum FileShareEnum As Integer
        None = &H0
        Read = &H1
        Write = &H2
        Delete = &H3
    End Enum

    Public Enum FileModeEnum As Integer
        [New] = 1
        CreateAlways = 2
        OpenExisting = 3
        OpenAlways = 4
        TruncateExisting = 5
    End Enum

    Public Enum FileAttributesEnum As Integer
        [Readonly] = &H1
        Hidden = &H2
        System = &H4
        Directory = &H10
        Archive = &H20
        Device = &H40
        Normal = &H80
        Temporary = &H100
        SparseFile = &H200
        ReparsePoint = &H400
        Compressed = &H800
        Offline = &H1000
        NotContentIndexed = &H2000
        Encrypted = &H4000
        Write_Through = &H80000000
        Overlapped = &H40000000
        NoBuffering = &H20000000
        RandomAccess = &H10000000
        SequentialScan = &H8000000
        DeleteOnClose = &H4000000
        BackupSemantics = &H2000000
        PosixSemantics = &H1000000
        OpenReparsePoint = &H200000
        OpenNoRecall = &H100000
        FirstPipeInstance = &H80000
    End Enum

    <DllImport("kernel32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Private Shared Function CreateFile( _
        ByVal fileName As String, _
        <MarshalAs(UnmanagedType.U4)> ByVal desiredAccess As Integer, _
        <MarshalAs(UnmanagedType.U4)> ByVal shareMode As Integer, _
        ByVal securityAttributes As IntPtr, _
        <MarshalAs(UnmanagedType.U4)> ByVal creationDisposition As Integer, _
        ByVal flags As Integer, _
        ByVal template As IntPtr) As SafeFileHandle
    End Function

    <DllImport("kernel32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function GetFileTime( _
        ByVal hFile As SafeFileHandle, _
        <Out()> ByRef lpCreationTime As ComTypes.FILETIME, _
        <Out()> ByRef lpLastAccessTime As ComTypes.FILETIME, _
        <Out()> ByRef lpLastWriteTime As ComTypes.FILETIME) As Boolean
    End Function

    <DllImport("kernel32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function SetFileTime( _
        ByVal hFile As SafeFileHandle, _
        <[In]()> ByRef lpCreationTime As ComTypes.FILETIME, _
        <[In]()> ByRef lpLastAccessedTime As ComTypes.FILETIME, _
        <[In]()> ByRef lpLastWriteTime As ComTypes.FILETIME) As Boolean
    End Function

    Public Shared Function GetFileHandle(ByVal pathname As String, ByVal access As Integer) As SafeFileHandle
        KernelApi.AddPrivilege(KernelApi.SE_BACKUP_NAME)
        KernelApi.AddPrivilege(KernelApi.SE_RESTORE_NAME)

        Dim ptr As SafeFileHandle
        ptr = CreateFile(pathname, _
                        access, _
                        0, _
                        IntPtr.Zero, _
                        OPEN_EXISTING, _
                        FILE_FLAG_BACKUP_SEMANTICS, _
                        IntPtr.Zero)
        If ptr.IsInvalid Then
            Throw New Win32Exception(Marshal.GetLastWin32Error())
        End If
        Return ptr
    End Function

    Public Shared Sub GetFileDate(ByVal pathname As String, ByRef create As DateTime, ByRef access As DateTime, ByRef write As DateTime)
        Dim cft As New ComTypes.FILETIME
        Dim aft As New ComTypes.FILETIME
        Dim wft As New ComTypes.FILETIME

        GetFileDate(pathname, cft, aft, wft)

        create = TimeApi.FileTimeToDateTime(cft)
        access = TimeApi.FileTimeToDateTime(aft)
        write = TimeApi.FileTimeToDateTime(wft)
    End Sub

    Public Shared Sub GetFileDate(ByVal pathname As String, ByRef create As ComTypes.FILETIME, ByRef access As ComTypes.FILETIME, ByRef write As ComTypes.FILETIME)
        Dim ptr As SafeFileHandle = Nothing
        Try
            ptr = FileApi.GetFileHandle(pathname, READ_CONTROL)
            FileApi.GetFileTime(ptr, create, access, write)
        Finally
            If ptr IsNot Nothing Then
                ptr.Close()
            End If
        End Try
    End Sub

    Public Shared Sub SetFileDate(ByVal pathname As String, ByVal create As DateTime, ByVal access As DateTime, ByVal write As DateTime)
        Dim hPtr As SafeFileHandle = Nothing
        Try
            hPtr = GetFileHandle(pathname, GENERIC_READ Or GENERIC_WRITE)
            SetFileDate(hPtr, create, access, write)
        Finally
            If hPtr IsNot Nothing Then
                hPtr.Close()
            End If
        End Try
    End Sub

    Public Shared Sub SetFileDate(ByVal hPtr As SafeFileHandle, ByVal create As DateTime, ByVal access As DateTime, ByVal write As DateTime)
        Dim cft As ComTypes.FILETIME = Nothing
        Dim aft As ComTypes.FILETIME = Nothing
        Dim wft As ComTypes.FILETIME = Nothing

        If create <> Nothing AndAlso create <> DateTime.MaxValue AndAlso create <> DateTime.MinValue Then
            cft = TimeApi.DateTimeToFileTime(create)
            If Not SetFileTime(hPtr, cft, Nothing, Nothing) Then
                Throw New Win32Exception(Marshal.GetLastWin32Error())
            End If
        End If
        If access <> Nothing AndAlso access <> DateTime.MaxValue AndAlso access <> DateTime.MinValue Then
            aft = TimeApi.DateTimeToFileTime(access)
            If Not SetFileTime(hPtr, Nothing, aft, Nothing) Then
                Throw New Win32Exception(Marshal.GetLastWin32Error())
            End If
        End If
        If write <> Nothing AndAlso write <> DateTime.MaxValue AndAlso write <> DateTime.MinValue Then
            wft = TimeApi.DateTimeToFileTime(write)
            If Not SetFileTime(hPtr, Nothing, Nothing, wft) Then
                Throw New Win32Exception(Marshal.GetLastWin32Error())
            End If
        End If
    End Sub

    Public Shared Function GetUniquePath(ByVal path As String) As String
        Dim info As New DirectoryInfo(path)
        Return GetUniquePath(info.Parent.FullName, info.Name)
    End Function

    Public Shared Function GetUniquePath(ByVal root As String, ByVal name As String) As String
        Dim cnt As Integer = 0
        Dim pathname As String
        pathname = Path.Combine(root, name)
        Do While Directory.Exists(pathname)
            cnt += 1
            pathname = Path.Combine(root, name + CStr(cnt))
        Loop
        Return pathname
    End Function

End Class
