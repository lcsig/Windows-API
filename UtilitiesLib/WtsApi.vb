Imports System.Runtime.InteropServices
Imports System.ComponentModel
Imports System.Text
Imports System.Diagnostics

Public Class WtsApi

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)> _
    Public Structure WTS_SESSION_INFO
        Public SessionId As Integer
        <MarshalAs(UnmanagedType.LPStr)> Public pWinStationName As String
        Public state As WTS_CONNECTSTATE_CLASS
    End Structure

    Public Class WtsSession
        Public SessionId As Integer
        Public StationName As String
        Public ConnectionState As WTS_CONNECTSTATE_CLASS
        Public UserDomain As String
        Public UserName As String
        Public UserSid As String
    End Class

    Public Enum WTS_CONNECTSTATE_CLASS
        WTSActive
        WTSConnected
        WTSConnectQuery
        WTSShadow
        WTSDisconnected
        WTSIdle
        WTSListen
        WTSReset
        WTSDown
        WTSInit
    End Enum

    Public Enum WTS_INFO_CLASS
        WTSInitialProgram = 0
        WTSApplicationName = 1
        WTSWorkingDirectory = 2
        WTSOEMId = 3
        WTSSessionId = 4
        WTSUserName = 5
        WTSWinStationName = 6
        WTSDomainName = 7
        WTSConnectState = 8
        WTSClientBuildNumber = 9
        WTSClientName = 10
        WTSClientDirectory = 11
        WTSClientProductId = 12
        WTSClientHardwareId = 13
        WTSClientAddress = 14
        WTSClientDisplay = 15
        WTSClientProtocolType = 16
        WTSIdleTime = 17
        WTSLogonTime = 18
        WTSIncomingBytes = 19
        WTSOutgoingBytes = 20
        WTSIncomingFrames = 21
        WTSOutgoingFrames = 22
    End Enum

    'WTSQueryUserToken errors
    Private Const ERROR_PRIVILEGE_NOT_HELD As Integer = 1314 'needs tcb privilege
    Private Const ERROR_INVALID_PARAMETER As Integer = 87
    Private Const ERROR_ACCESS_DENIED As Integer = 5 'has tcb privilege, but it must be running as LocalSystem
    Private Const ERROR_CTX_WINSTATION_NOT_FOUND As Integer = 7022
    Private Const ERROR_NO_TOKEN As Integer = 1008

    Private Const WTS_NO_SESSION As Integer = &HFFFFFFFF
    Private Const WTS_CURRENT_SERVER_HANDLE As Long = 0
    Private Const WTS_CURRENT_SESSION As Long = -1

    <DllImport("wtsapi32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Private Shared Function WTSEnumerateSessions( _
        ByVal hServer As IntPtr, _
        <MarshalAs(UnmanagedType.U4)> ByVal Reserved As Integer, _
        <MarshalAs(UnmanagedType.U4)> ByVal Version As Integer, _
        ByRef ppSessionInfo As IntPtr, _
        <MarshalAs(UnmanagedType.U4)> ByRef pCount As Integer) As Integer
    End Function

    <DllImport("wtsapi32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Private Shared Function WTSQuerySessionInformation( _
         ByVal hServer As Integer, _
         ByVal SessionId As Integer, _
         ByVal InfoClass As WTS_INFO_CLASS, _
         ByRef ppBuffer As IntPtr, _
         ByRef pCount As Int32) As Integer
    End Function

    <DllImport("wtsapi32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Private Shared Function WTSFreeMemory( _
        ByVal memory As IntPtr) As Integer
    End Function

    <DllImport("wtsapi32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Private Shared Function WTSOpenServer( _
        ByVal pServerName As String) As IntPtr
    End Function

    <DllImport("wtsapi32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Private Shared Sub WTSCloseServer( _
        ByVal hServer As IntPtr)
    End Sub

    <DllImport("kernel32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function WTSGetActiveConsoleSessionId() As Integer
    End Function

    <DllImport("Wtsapi32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Private Shared Function WTSDisconnectSession( _
        ByVal hServer As Integer, _
        ByVal SessionID As Integer, _
        ByVal bWait As Boolean) As Integer
    End Function

    <DllImport("Wtsapi32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Private Shared Function WTSLogoffSession( _
        ByVal hServer As Integer, _
        ByVal Sessionid As Integer, _
        ByVal bWait As Boolean) As Integer
    End Function

    <DllImport("Wtsapi32.dll", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Shared Function WTSQueryUserToken( _
        ByVal Sessionid As Integer, _
        ByRef pHandle As IntPtr) As Integer
    End Function

    Public Shared Function WaitForSessionRelease(ByVal username As String, ByVal milliseconds As Integer) As Boolean
        Dim began As DateTime = Now

        While Now <= began.AddMilliseconds(milliseconds)
            If WtsApi.UserIsLoggedIn(username) Then
                Logger.WriteEntry("Waiting for session to release: " + username, EventLogEntryType.Warning)
                System.Threading.Thread.Sleep(milliseconds)
            Else
                Return True
            End If
        End While

        Return False
    End Function

    'Public Shared Function WaitForSessionRelease(ByVal sessionId As Integer, ByVal milliseconds As Integer) As Boolean
    '    Dim began As DateTime = Now

    '    While Now < began.AddMilliseconds(milliseconds)
    '        Try
    '            If Not String.IsNullOrEmpty(GetSessionInfo(sessionId, WTS_INFO_CLASS.WTSConnectState)) Then
    '                'continue
    '            End If
    '        Catch ex As Exception
    '            'session went out of scope
    '            Return True
    '        End Try
    '        System.Threading.Thread.Sleep(500)
    '    End While

    '    Return False
    'End Function

    Public Shared Function UserIsLoggedIn(ByVal username As String) As Boolean
        Try
            If WtsApi.GetSession(username) Is Nothing Then
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            Return True
        End Try
    End Function

    Public Shared Function GetUserToken(ByVal sessionID As Integer) As IntPtr
        Dim result As Integer
        Dim token As IntPtr = IntPtr.Zero

        KernelApi.AddPrivilege(KernelApi.SE_TCB_PRIVILEGE)

        result = WTSQueryUserToken(sessionID, token)
        If result = 0 Then
            Throw New Win32Exception(Marshal.GetLastWin32Error())
        End If
        Return token
    End Function

    Public Shared Function GetUserName(ByVal sessionID As Integer) As String
        Return GetSessionInfo(sessionID, WTS_INFO_CLASS.WTSUserName)
    End Function

    Public Shared Function GetUserDomain(ByVal sessionID As Integer) As String
        Return GetSessionInfo(sessionID, WTS_INFO_CLASS.WTSDomainName)
    End Function

    Public Shared Function GetUserNameFull(ByVal sessionid As Integer) As String
        Return GetUserDomain(sessionid).ToUpper() + "\" + GetUserName(sessionid)
    End Function

    Public Shared Function GetCurrentSessionId() As Integer
        Return CInt(GetSessionInfo(WTS_CURRENT_SESSION, WTS_INFO_CLASS.WTSSessionId))
    End Function

    Public Shared Function GetCurrentSessionInfo(ByVal info As WTS_INFO_CLASS) As String
        Return GetSessionInfo(WTS_CURRENT_SESSION, info)
    End Function

    Public Shared Function GetSessionInfo(ByVal sessionID As Integer, ByVal info As WTS_INFO_CLASS) As String
        Dim ret As String = ""
        Dim ppBuffer As IntPtr = IntPtr.Zero
        Dim bufSize As Integer
        Dim result As Integer
        Try
            result = WTSQuerySessionInformation(WTS_CURRENT_SERVER_HANDLE, sessionID, info, ppBuffer, bufSize)
            If result = 0 Then
                Throw New Win32Exception(Marshal.GetLastWin32Error())
            End If

            Select Case info
                Case WTS_INFO_CLASS.WTSApplicationName
                    ret = Marshal.PtrToStringAuto(ppBuffer)
                Case WTS_INFO_CLASS.WTSClientName
                    ret = Marshal.PtrToStringAuto(ppBuffer)
                Case WTS_INFO_CLASS.WTSDomainName
                    ret = Marshal.PtrToStringAuto(ppBuffer)
                Case WTS_INFO_CLASS.WTSInitialProgram
                    ret = Marshal.PtrToStringAuto(ppBuffer)
                Case WTS_INFO_CLASS.WTSUserName
                    ret = Marshal.PtrToStringAuto(ppBuffer)
                Case WTS_INFO_CLASS.WTSWinStationName
                    ret = Marshal.PtrToStringAuto(ppBuffer)
                Case WTS_INFO_CLASS.WTSWorkingDirectory
                    ret = Marshal.PtrToStringAuto(ppBuffer)
                Case WTS_INFO_CLASS.WTSConnectState
                    ret = CType(Marshal.ReadInt32(ppBuffer), WTS_CONNECTSTATE_CLASS).ToString()
                Case WTS_INFO_CLASS.WTSSessionId
                    ret = Convert.ToString(Marshal.ReadInt32(ppBuffer))
                Case Else
                    Throw New ArgumentOutOfRangeException(info.ToString())
            End Select
            Return ret
        Finally
            WTSFreeMemory(ppBuffer)
        End Try
    End Function

    Public Shared Function GetSession(ByVal username As String) As WtsSession
        username = Utilities.GetUserName(username)

        For Each s As WtsApi.WtsSession In WtsApi.GetSessions()
            If Utilities.GetUserName(s.UserName).ToLower() = username.ToLower() Then
                Return s
            End If
        Next
        Return Nothing
    End Function

    Public Shared Function GetActiveSession() As WtsSession
        Dim retries As Integer = 0
        Try
Retry:
            Dim sessions() As WtsApi.WtsSession = WtsApi.GetSessions()

            'priority to WtsActive sessions
            For Each s As WtsApi.WtsSession In sessions
                If s.ConnectionState = WTS_CONNECTSTATE_CLASS.WTSActive Then
                    Return s
                End If
            Next
            For Each s As WtsApi.WtsSession In sessions
                If s.ConnectionState = WTS_CONNECTSTATE_CLASS.WTSConnected Then
                    Return s
                End If
            Next
        Catch ex As Exception
            If retries < 3 Then
                retries += 1
                Threading.Thread.Sleep(500)
                GoTo Retry
            End If
            Logger.WriteEntry("Error retrieving active session: " + ex.Message, EventLogEntryType.Error)
        End Try
        Return Nothing
    End Function

    Public Shared Function ActiveSessionIsLoggedIn() As Boolean
        Dim s As WtsApi.WtsSession
        s = GetActiveSession()
        If s Is Nothing Then
            Return False
        ElseIf s.UserName = "" Then
            Return False
        Else
            Return True
        End If
    End Function

    Public Shared Function GetSession(ByVal sessionId As Integer) As WtsSession
        For Each s As WtsSession In WtsApi.GetSessions()
            If s.SessionId = sessionId Then
                Return s
            End If
        Next
        Return Nothing
    End Function

    Public Shared Function GetSessions() As WtsSession()
        Dim ppSessionInfo As IntPtr = IntPtr.Zero
        Dim result As Integer
        Dim sessionCount As Integer
        Dim structSize As Integer
        Dim structIndexer As Integer

        Try
            result = WTSEnumerateSessions(IntPtr.Zero, 0, 1, ppSessionInfo, sessionCount)
            If result = 0 Then
                Throw New Win32Exception(Marshal.GetLastWin32Error())
            End If

            structSize = Marshal.SizeOf(GetType(WTS_SESSION_INFO))
            structIndexer = CInt(ppSessionInfo)

            Dim sessions(sessionCount - 1) As WtsSession

            For i As Integer = 0 To sessionCount - 1
                Dim si As WTS_SESSION_INFO
                si = CType(Marshal.PtrToStructure(CType(structIndexer, IntPtr), GetType(WTS_SESSION_INFO)), WTS_SESSION_INFO)

                sessions(i) = New WtsSession()
                sessions(i).SessionId = si.SessionId
                sessions(i).ConnectionState = si.state
                sessions(i).StationName = si.pWinStationName
                Dim domain As String = GetSessionInfo(si.SessionId, WTS_INFO_CLASS.WTSDomainName)
                Dim user As String = GetSessionInfo(si.SessionId, WTS_INFO_CLASS.WTSUserName)
                sessions(i).UserDomain = domain
                sessions(i).UserName = user
                sessions(i).UserSid = KernelApi.GetSID(domain + "\" + user)

                structIndexer += structSize
            Next
            Return sessions
        Catch ex As Exception
            Throw New ApplicationException("Error enumerating WtsSessions: " + ex.Message)
        Finally
            WTSFreeMemory(ppSessionInfo)
        End Try
    End Function

    Public Shared Sub LogoffSession(ByVal sessionId As Integer, ByVal synchronous As Boolean)
        Dim result As Integer
        result = WtsApi.WTSLogoffSession(WTS_CURRENT_SERVER_HANDLE, sessionId, synchronous)
        If result = 0 Then
            Throw New Win32Exception(Marshal.GetLastWin32Error())
        End If
    End Sub

    Public Shared Sub DisconnectSession(ByVal sessionId As Integer, ByVal synchronous As Boolean)
        Dim result As Integer
        result = WtsApi.WTSDisconnectSession(WTS_CURRENT_SERVER_HANDLE, sessionId, synchronous)
        If result = 0 Then
            Throw New Win32Exception(Marshal.GetLastWin32Error())
        End If
    End Sub

    Public Shared Function AllUsersLoggedOff() As Boolean
        For Each s As WtsSession In WtsApi.GetSessions()
            If s.UserName <> "" Then Return False
        Next
        Return True
    End Function

    Public Shared Function WaitForConnectedSession(ByVal timeoutMilliseconds As Integer) As WtsSession
        Dim session As WtsSession = Nothing
        Dim elapsed As Integer = 0

        While elapsed < timeoutMilliseconds
            Try
                For Each s As WtsSession In WtsApi.GetSessions()
                    If s.ConnectionState = WTS_CONNECTSTATE_CLASS.WTSConnected Then
                        session = s
                        Exit While
                    End If
                Next
            Catch ex As Exception
                'ignore
            End Try
            Threading.Thread.Sleep(500)
            elapsed += 500
        End While

        Return session
    End Function

End Class
